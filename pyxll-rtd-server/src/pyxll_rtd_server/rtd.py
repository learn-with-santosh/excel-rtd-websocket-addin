"""PyXLL RTD topics backed by the repository's WebSocket server."""

from __future__ import annotations

import json
import threading
from typing import Any

import websocket
from pyxll import RTD, xl_func

_DEFAULT_URL = "ws://localhost:8080"
_SUPPORTED_FIELDS = {"LAST", "OPEN", "HIGH", "LOW", "VOLUME"}


class WebSocketQuoteRTD(RTD):
    """Stream one quote field for one symbol into an Excel cell."""

    def __init__(self, symbol: str, field: str, url: str = _DEFAULT_URL):
        super().__init__()
        self.symbol = symbol.upper().strip()
        self.field = field.upper().strip()
        self.url = url
        self._socket: websocket.WebSocketApp | None = None
        self._thread: threading.Thread | None = None
        self._stop = threading.Event()

    def connect(self) -> None:
        self._stop.clear()
        self._thread = threading.Thread(
            target=self._run,
            name=f"pyxll-rtd-{self.symbol}-{self.field}",
            daemon=True,
        )
        self._thread.start()

    def disconnect(self) -> None:
        self._stop.set()
        if self._socket is not None:
            self._socket.close()
        if self._thread is not None and self._thread is not threading.current_thread():
            self._thread.join(timeout=2)

    def _run(self) -> None:
        def on_open(socket: websocket.WebSocketApp) -> None:
            socket.send(json.dumps({"action": "subscribe", "symbols": [self.symbol]}))

        def on_message(_: websocket.WebSocketApp, message: str) -> None:
            try:
                payload: dict[str, Any] = json.loads(message)
                value = payload.get("data", {}).get(self.symbol, {}).get(self.field)
            except (TypeError, ValueError, AttributeError):
                return
            if value is not None:
                self.set_value(value)

        while not self._stop.is_set():
            self._socket = websocket.WebSocketApp(
                self.url,
                on_open=on_open,
                on_message=on_message,
            )
            self._socket.run_forever()
            self._socket = None
            if not self._stop.wait(1):
                continue


@xl_func("string symbol, string field, string url: var")
def websocket_rtd(symbol: str, field: str, url: str = _DEFAULT_URL) -> WebSocketQuoteRTD:
    """Return an RTD topic for a symbol and market-data field."""
    normalized_field = field.upper().strip()
    if not symbol.strip():
        raise ValueError("symbol must not be empty")
    if normalized_field not in _SUPPORTED_FIELDS:
        supported = ", ".join(sorted(_SUPPORTED_FIELDS))
        raise ValueError(f"field must be one of: {supported}")
    return WebSocketQuoteRTD(symbol, normalized_field, url.strip() or _DEFAULT_URL)

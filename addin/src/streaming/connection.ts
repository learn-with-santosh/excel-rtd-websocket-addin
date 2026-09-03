/* global WebSocket, console, setTimeout */

export type QuoteUpdate = Record<string, number>;

type DataMessage =
  | { type: "snapshot"; data: Record<string, QuoteUpdate> }
  | { type: "tick"; data: Record<string, QuoteUpdate> }
  | { type: "error"; code: string; message: string };

type DataHandler = (
  kind: "snapshot" | "tick",
  data: Record<string, QuoteUpdate>
) => void;

// TODO: move to config/env before deploying
const WS_URL = "ws://localhost:8080";
const BASE_BACKOFF_MS = 1000;
const MAX_BACKOFF_MS = 15000;

let socket: WebSocket | null = null;
let attempt = 0;
let reconnectTimer: ReturnType<typeof setTimeout> | null = null;

const desiredSymbols = new Set<string>();
let dataHandler: DataHandler | null = null;

/** Registry registers a single handler for all incoming data. */
export function onData(handler: DataHandler): void {
  dataHandler = handler;
}

/** Called when a symbol's reference count goes 0 → 1. */
export function acquire(symbol: string): void {
  const isNew = !desiredSymbols.has(symbol);
  desiredSymbols.add(symbol);
  if (!isNew) return;

  ensureConnected();
  if (isOpen()) send({ action: "subscribe", symbols: [symbol] });
  // If the socket is still connecting, onopen will re-subscribe everything.
}

/** Called when a symbol's reference count goes 1 → 0. */
export function release(symbol: string): void {
  if (!desiredSymbols.delete(symbol)) return;
  if (isOpen()) send({ action: "unsubscribe", symbols: [symbol] });
}

function isOpen(): boolean {
  return socket !== null && socket.readyState === WebSocket.OPEN;
}

function send(obj: unknown): void {
  if (isOpen()) socket!.send(JSON.stringify(obj));
}

function ensureConnected(): void {
  if (socket || desiredSymbols.size === 0) return;
  connect();
}

function connect(): void {
  socket = new WebSocket(WS_URL);

  socket.onopen = () => {
    attempt = 0;
    if (desiredSymbols.size > 0) {
      send({ action: "subscribe", symbols: Array.from(desiredSymbols) });
    }
  };

  socket.onmessage = (event: MessageEvent) => {
    let msg: DataMessage;
    try {
      msg = JSON.parse(event.data as string) as DataMessage;
    } catch {
      return; // ignore malformed frames
    }
    if ((msg.type === "snapshot" || msg.type === "tick") && dataHandler) {
      dataHandler(msg.type, msg.data);
    } else if (msg.type === "error") {
      console.warn("mock-ws-server error:", msg.code, msg.message);
    }
  };

  socket.onclose = () => {
    socket = null;
    scheduleReconnect();
  };

  socket.onerror = () => {
    // onclose always follows; nothing to do here.
  };
}

function scheduleReconnect(): void {
  if (desiredSymbols.size === 0 || reconnectTimer) return;
  const delay = Math.min(BASE_BACKOFF_MS * Math.pow(2, attempt), MAX_BACKOFF_MS);
  attempt += 1;
  reconnectTimer = setTimeout(() => {
    reconnectTimer = null;
    if (desiredSymbols.size > 0 && !socket) connect();
  }, delay);
}
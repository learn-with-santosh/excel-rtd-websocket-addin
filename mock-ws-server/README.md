# Mock WebSocket Market Data Server

A tiny Node.js WebSocket server that simulates a live market data feed, built for testing the [Excel RTD WebSocket Add-in](../addin). It stands in for a real data provider so you can develop and test streaming custom functions without a broker account or paid feed.

Part of the [`excel-rtd-websocket-addin`](https://github.com/<your-username>/excel-rtd-websocket-addin) monorepo — but perfectly usable standalone with any WebSocket client.

## Features

- Single file, single dependency (`ws`)
- Accepts **any symbol** — mock data is generated on demand
- Snapshot + streaming ticks over one socket
- Partial ticks (only changed fields), like real market feeds
- Per-socket subscription tracking with fan-out
- Deterministic base prices derived from the symbol name — stable across restarts
- No auth, no TLS, no persistence — intentionally simple

## Requirements

- Node.js 18+
- npm

## Installation

```bash
cd mock-ws-server
npm install
```

## Running

```bash
node server.js
# Mock server listening on ws://localhost:8080
```

Configuration via environment variables:

| Variable | Default | Description |
|---|---|---|
| `PORT` | `8080` | Listening port |
| `TICK_INTERVAL_MS` | `1000` | Tick generation interval per symbol |

```bash
PORT=9090 TICK_INTERVAL_MS=500 node server.js
```

## Wire Protocol

All messages are UTF-8 JSON, one object per WebSocket message.

### Client → Server

**Subscribe**

```json
{"action": "subscribe", "symbols": ["ACC.NS", "RELIANCE.NS"]}
```

**Unsubscribe**

```json
{"action": "unsubscribe", "symbols": ["ACC.NS"]}
```

| Field | Type | Notes |
|---|---|---|
| `action` | string | `subscribe` or `unsubscribe` |
| `symbols` | string[] | One or more symbols; letters, digits, `.`, `-` |

### Server → Client

**Snapshot** — sent immediately in response to every `subscribe`:

```json
{
  "type": "snapshot",
  "data": {
    "ACC.NS": { "LAST": 1024.5, "OPEN": 1010, "HIGH": 1030, "LOW": 1005, "VOLUME": 123456 }
  }
}
```

**Tick** — pushed every interval per active symbol; contains **only changed fields**:

```json
{"type": "tick", "data": {"ACC.NS": {"LAST": 1025.1, "VOLUME": 124000}}}
```

**Error** — sent when an incoming message is malformed:

```json
{"type": "error", "code": "INVALID_REQUEST", "message": "symbols must be a non-empty array"}
```

## Fields

| Field | Type | Updated on tick? |
|---|---|---|
| `LAST` | number | Every tick (random walk, ±0.5%) |
| `VOLUME` | integer | Every tick (monotonic increase) |
| `HIGH` | number | Only when `LAST` exceeds it |
| `LOW` | number | Only when `LAST` drops below it |
| `OPEN` | number | Never — set once in the snapshot |

## Behavior

- **Snapshot on subscribe.** Every subscribe receives an immediate full snapshot of all requested symbols — no separate REST call needed.
- **Any symbol works.** Base price is derived deterministically from the symbol name, so `ACC.NS` starts at the same price on every restart — handy for repeatable tests.
- **Fan-out.** Multiple sockets can subscribe to the same symbol; each receives its own tick stream.
- **Unsubscribe is per-socket.** Stopping ticks for one socket doesn't affect others still subscribed.
- **Disconnect cleanup.** All subscriptions held by a socket are dropped when it closes.
- **No heartbeat.** The mock relies on TCP close; the `ws` library answers protocol-level pings automatically.

## Quick Test with wscat

```bash
npm install -g wscat
wscat -c ws://localhost:8080
```

```
> {"action":"subscribe","symbols":["ACC.NS"]}

< {"type":"snapshot","data":{"ACC.NS":{"LAST":1024.5,"OPEN":1010,"HIGH":1030,"LOW":1005,"VOLUME":123456}}}
< {"type":"tick","data":{"ACC.NS":{"LAST":1027.2,"VOLUME":124000}}}
< {"type":"tick","data":{"ACC.NS":{"LAST":1026.8,"VOLUME":124100}}}
```

You should see a snapshot instantly, then one tick per second.

## Using with the Add-in

1. **Terminal 1** — add-in: `cd addin && npm run dev-server`
2. **Terminal 2** — this server: `cd mock-ws-server && node server.js`
3. Sideload the add-in, then in Excel enter:
   `=GETLIVEDATA("ACC.NS","LAST,OPEN,HIGH,LOW,VOLUME")`
4. Expect an instant snapshot, then live updates every second.

**Reconnection test:** kill this server while cells are streaming, restart it, and confirm the cells recover within the add-in's backoff window.

## Troubleshooting

| Symptom | Cause / Fix |
|---|---|
| `EADDRINUSE` on start | Port 8080 already taken — set `PORT` |
| Add-in never connects on macOS | `https://` page → `ws://localhost` is blocked by Safari-based hosts; use `wss` or test on Windows/WebView2 |
| Snapshot arrives but no ticks | Client likely never sent subscribe; malformed JSON returns an `error` message |
| `wscat: command not found` | Global npm bin not on your `PATH` |

## Project Structure

```
mock-ws-server/
├── server.js     # the entire server (~100 lines)
├── package.json  # one dependency: ws
└── README.md
```

## Security Note

This server has **no authentication, no TLS, and no rate limiting**. It is for local development only — never expose it to the public internet.

## Roadmap

- `wss://` support using the self-signed certs office-addin-dev-certs already generates
- Simulated latency and random disconnects (to exercise client reconnect logic)
- Optional token auth
- Record/replay mode for repeatable test data

## License

MIT — same as the repository.


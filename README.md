# Excel RTD WebSocket Add-in

An Excel custom-functions add-in that demonstrates regular functions, spill ranges, streaming functions, and live market data from a local WebSocket server.

## Requirements

- Node.js 18 or later
- npm
- Microsoft Excel with Office Add-ins support

## Project Structure

```text
addin/          Excel add-in source, manifest, and HTTPS development server
mock-ws-server/ Local WebSocket market-data server
run-applications.bat  Starts Excel debugging and the mock server
```

## Run Everything

From the repository root, double-click:

```text
run-applications.bat
```

This starts the add-in in Excel debug mode and starts the mock WebSocket server.

To start the pieces manually, use separate terminals:

```powershell
cd addin
npm install
npm run start
```

```powershell
cd mock-ws-server
npm install
npm start
```

The add-in development server is available at `https://localhost:3000` and the mock WebSocket server listens at `ws://localhost:8080`.

## Excel Namespace

The namespace is `COG`, configured in `addin/manifest.xml`. Use `COG.` before every custom function name.

## Functions

### ADD

Adds two numbers.

```excel
=COG.ADD(10, 5)
```

Output:

```text
15
```

### GREET

Returns a greeting.

```excel
=COG.GREET("dfad")
```

Output:

```text
Hello, dfad!
```

### SAMPLESPILL

Returns a two-row table that spills into six cells across and two rows down.

```excel
=COG.SAMPLESPILL()
```

Output:

| SYMBOL | LAST | OPEN | HIGH | LOW | VOLUME |
|---|---:|---:|---:|---:|---:|
| ACC.NS | 1024.5 | 1010 | 1030 | 1005 | 123456 |

### CLOCK

Streams the current local time and updates once per second.

```excel
=COG.CLOCK()
```

Example output:

```text
14:32:01
14:32:02
14:32:03
```

### INCREMENT

Streams an increasing number once per second. The value increases by the supplied amount.

```excel
=COG.INCREMENT(5)
```

Example output:

```text
5
10
15
```

### LOG

Writes the message to the browser console and returns the same message to Excel.

```excel
=COG.LOG("Test message")
```

Excel output:

```text
Test message
```

### TESTSTREAM

Streams a count that starts at the supplied value and increases by one per second.

```excel
=COG.TESTSTREAM(100)
```

Example output:

```text
101
102
103
```

### GETLIVEDATA

Connects to the local mock WebSocket server and streams selected market fields. The result spills as a labeled table.

```excel
=COG.GETLIVEDATA("ACC.NS","LAST,OPEN,HIGH,LOW,VOLUME")
```

Initial layout:

| SYMBOL | LAST | OPEN | HIGH | LOW | VOLUME |
|---|---:|---:|---:|---:|---:|
| ACC.NS | Connecting... | Connecting... | Connecting... | Connecting... | Connecting... |

Example live output:

| SYMBOL | LAST | OPEN | HIGH | LOW | VOLUME |
|---|---:|---:|---:|---:|---:|
| ACC.NS | 1025.10 | 1010.00 | 1030.00 | 1005.00 | 124000 |

Supported fields are `LAST`, `OPEN`, `HIGH`, `LOW`, and `VOLUME`. Field names are case-insensitive and duplicate fields are removed.

## Build

Build the add-in and regenerate `addin/dist/functions.js` and `addin/dist/functions.json`:

```powershell
npm --prefix addin run build
```

## Troubleshooting

- If functions do not appear, close Excel, run `npm --prefix addin run stop`, then run `npm --prefix addin run start` and reload the manifest.
- Use the current namespace: `=COG.TESTSTREAM(1)`.
- `GETLIVEDATA` requires the mock server to be running on port `8080`.
- If port `3000` or `8080` is busy, stop the existing process or change the corresponding configuration.
- The add-in uses HTTPS locally. Accept/install the development certificate when prompted.

## License

MIT

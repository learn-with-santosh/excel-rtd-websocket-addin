/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office, WebSocket */

type Quote = Record<string, number>;
type DataMessage =
  | { type: "snapshot"; data: Record<string, Quote> }
  | { type: "tick"; data: Record<string, Quote> };

const WS_URL = "ws://localhost:8080";
const validFields = ["LAST", "OPEN", "HIGH", "LOW", "VOLUME"];
let socket: WebSocket | null = null;
let activeSymbols: string[] = [];
let activeFields: string[] = [];
const quotes = new Map<string, Quote>();

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg")!.style.display = "none";
  document.getElementById("app-body")!.style.display = "flex";
  document.getElementById("connect")!.addEventListener("click", connectBoard);
  document.getElementById("copy")!.addEventListener("click", copyTable);
  document.getElementById("paste")!.addEventListener("click", pasteToExcel);
  renderTable();
});

function parseList(value: string): string[] {
  return value
    .split(/[\n,]+/)
    .map((item) => item.trim().toUpperCase())
    .filter((item, index, items) => item.length > 0 && items.indexOf(item) === index);
}

function connectBoard(): void {
  const symbols = parseList((document.getElementById("symbols") as HTMLTextAreaElement).value);
  const fields = parseList((document.getElementById("fields") as HTMLInputElement).value);
  const invalid = fields.filter((field) => !validFields.includes(field));

  if (symbols.length === 0) {
    setStatus("Add at least one symbol.", "error");
    return;
  }
  if (invalid.length > 0) {
    setStatus(`Unknown field: ${invalid[0]}.`, "error");
    return;
  }
  if (fields.length === 0) {
    setStatus("Add at least one field.", "error");
    return;
  }

  activeSymbols = symbols;
  activeFields = fields;
  quotes.clear();
  closeSocket();
  renderTable();
  setStatus(`Connecting to ${symbols.length} symbol${symbols.length === 1 ? "" : "s"}...`, "connecting");

  socket = new WebSocket(WS_URL);
  socket.onopen = () => {
    setConnection("Live");
    setStatus(`One stream for ${activeSymbols.length} unique symbol${activeSymbols.length === 1 ? "" : "s"}.`, "live");
    socket!.send(JSON.stringify({ action: "subscribe", symbols: activeSymbols }));
  };
  socket.onmessage = (event: MessageEvent) => {
    try {
      const message = JSON.parse(event.data as string) as DataMessage;
      if (message.type !== "snapshot" && message.type !== "tick") return;
      Object.entries(message.data).forEach(([symbol, update]) => {
        quotes.set(symbol, { ...(quotes.get(symbol) ?? {}), ...update });
      });
      renderTable();
    } catch {
      setStatus("Received an unreadable server update.", "error");
    }
  };
  socket.onerror = () => setStatus("WebSocket unavailable. Start the mock server and reconnect.", "error");
  socket.onclose = () => {
    setConnection("Offline");
    if (socket) setStatus("Connection closed.", "error");
    socket = null;
  };
}

function closeSocket(): void {
  if (!socket) return;
  socket.onclose = null;
  socket.close();
  socket = null;
  setConnection("Offline");
}

function renderTable(): void {
  const table = document.getElementById("quote-table")!;
  table.querySelector("thead")!.innerHTML = `<tr><th>Symbol</th>${activeFields.map((field) => `<th>${field}</th>`).join("")}</tr>`;
  table.querySelector("tbody")!.innerHTML = activeSymbols.map((symbol) => {
    const quote = quotes.get(symbol);
    return `<tr><th scope="row">${symbol}</th>${activeFields.map((field) => `<td>${quote?.[field] === undefined ? "..." : quote[field]}</td>`).join("")}</tr>`;
  }).join("");
  const enabled = activeSymbols.length > 0 && activeFields.length > 0;
  (document.getElementById("copy") as HTMLButtonElement).disabled = !enabled;
  (document.getElementById("paste") as HTMLButtonElement).disabled = !enabled;
}

function tableText(): string {
  return [
    ["Symbol", ...activeFields],
    ...activeSymbols.map((symbol) => [symbol, ...activeFields.map((field) => createFormula(symbol, field))]),
  ].map((row) => row.join("\t")).join("\n");
}

function createFormula(symbol: string, field: string): string {
  const escapedSymbol = symbol.replace(/"/g, "\"\"");
  return `=COG.GETLIVEDATA("${escapedSymbol}","${field}")`;
}

async function copyTable(): Promise<void> {
  try {
    await navigator.clipboard.writeText(tableText());
    setStatus("Copied live formulas. Paste them into Excel.", "live");
  } catch {
    setStatus("Clipboard access was blocked by the host.", "error");
  }
}

async function pasteToExcel(): Promise<void> {
  try {
    await Excel.run(async (context: Excel.RequestContext) => {
      const startCell = context.workbook.getSelectedRange().getCell(0, 0);
      const range = startCell.getResizedRange(activeSymbols.length, activeFields.length);
      range.formulas = [
        ["Symbol", ...activeFields],
        ...activeSymbols.map((symbol) => [symbol, ...activeFields.map((field) => createFormula(symbol, field))]),
      ];
      await context.sync();
    });
    setStatus("Live formulas pasted into the selected range.", "live");
  } catch (error) {
    console.error(error);
    setStatus("Select a starting cell in Excel, then try again.", "error");
  }
}

function setStatus(message: string, state: "connecting" | "live" | "error"): void {
  const status = document.getElementById("stream-status")!;
  status.textContent = message;
  status.dataset.state = state;
}

function setConnection(label: string): void {
  document.getElementById("connection-label")!.textContent = label;
}

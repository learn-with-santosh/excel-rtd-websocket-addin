/* global setTimeout, clearTimeout */

import { acquire, onData, release } from "./connection";

export type MergedQuote = Record<string, number | undefined>;
type QuoteListener = (quote: MergedQuote) => void;

const listeners = new Map<string, Set<QuoteListener>>();
const lastQuote = new Map<string, MergedQuote>();
const pendingFlush = new Map<string, ReturnType<typeof setTimeout>>();

const TICK_FLUSH_MS = 250; // coalesce ticks; snapshots bypass this

let initialized = false;

function ensureInit(): void {
  if (initialized) return;
  initialized = true;

  onData((kind, data) => {
    for (const symbol of Object.keys(data)) {
      const update = data[symbol];
      const merged: MergedQuote =
        kind === "snapshot"
          ? { ...update } // snapshot is a full replace → immediate first paint
          : { ...(lastQuote.get(symbol) ?? {}), ...update }; // tick = merge
      lastQuote.set(symbol, merged);

      if (kind === "snapshot") fanout(symbol);
      else scheduleFanout(symbol);
    }
  });
}

function fanout(symbol: string): void {
  const quote = lastQuote.get(symbol);
  const set = listeners.get(symbol);
  if (!quote || !set) return;
  for (const fn of set) {
    try {
      fn(quote);
    } catch {
      // one failing cell must not starve the others
    }
  }
}

function scheduleFanout(symbol: string): void {
  if (pendingFlush.has(symbol)) return;
  pendingFlush.set(
    symbol,
    setTimeout(() => {
      pendingFlush.delete(symbol);
      fanout(symbol);
    }, TICK_FLUSH_MS)
  );
}

export function subscribeQuote(symbol: string, listener: QuoteListener): void {
  ensureInit();

  let set = listeners.get(symbol);
  if (!set) {
    set = new Set();
    listeners.set(symbol, set);
  }
  set.add(listener);

  acquire(symbol);

  // If we already have data (e.g. same symbol in other cells), paint instantly.
  const cached = lastQuote.get(symbol);
  if (cached) listener(cached);
}

export function unsubscribeQuote(symbol: string, listener: QuoteListener): void {
  const set = listeners.get(symbol);
  if (!set) return;

  set.delete(listener);
  if (set.size === 0) {
    listeners.delete(symbol);
    lastQuote.delete(symbol);

    const t = pendingFlush.get(symbol);
    if (t) {
      clearTimeout(t);
      pendingFlush.delete(symbol);
    }

    release(symbol); // ref count hit zero → tell the server
  }
}
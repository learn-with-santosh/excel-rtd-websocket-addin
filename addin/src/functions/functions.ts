/* global clearInterval, console, CustomFunctions, setInterval */

import { parseFields } from "../streaming/fields";
import { subscribeQuote, unsubscribeQuote } from "../streaming/registry";

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Greets a person by name.
 * @customfunction
 * @param name Name of the person to greet.
 * @returns A greeting for the specified person.
 */
export function greet(name: string): string {
  return `Hello, ${name}!`;
}

/**
 * Returns sample data that spills into multiple cells.
 * @customfunction SAMPLESPILL
 * @returns {any[][]} A sample table with headers and values.
 */
export function sampleSpill(): string[][] {
  return [
    ["SYMBOL", "LAST", "OPEN", "HIGH", "LOW", "VOLUME"],
    ["ACC.NS", "1024.5", "1010", "1030", "1005", "123456"],
  ];
}


/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(
  incrementBy: number,
  invocation: CustomFunctions.StreamingInvocation<number>
): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}

/**
 * Streams live market data for a symbol with a header row.
 * Sample: =GETLIVEDATA("ACC.NS","LAST,OPEN,HIGH,LOW,VOLUME")
 * @customfunction GETLIVEDATA
 * @param {string} symbol Ticker symbol, e.g. ACC.NS
 * @param {string} fields Comma-separated: LAST, OPEN, HIGH, LOW, VOLUME
 * @param {CustomFunctions.StreamingInvocation<string[][]>} invocation Invocation
 */
export function getLiveData(
  symbol: string,
  fields: string,
  invocation: CustomFunctions.StreamingInvocation<string[][]>
): void {
  const trimmedSymbol = (symbol ?? "").trim().toUpperCase();

  if (!trimmedSymbol) {
    invocation.setResult(
      new CustomFunctions.Error(
        CustomFunctions.ErrorCode.invalidValue,
        'Symbol is required, e.g. "ACC.NS".'
      )
    );
    return;
  }

  let requested: string[];
  try {
    requested = parseFields(fields);
  } catch (e) {
    invocation.setResult(
      new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, (e as Error).message)
    );
    return;
  }

  const header: string[] = ["SYMBOL", ...requested];
  let canceled = false;

  // Rectangular: same width as header
  invocation.setResult([
    header,
    [trimmedSymbol, ...requested.map(() => "Connecting…")],
  ]);

  const listener = (quote: Record<string, number | undefined>): void => {
    if (canceled) return;
    const values = requested.map((f) => {
      const v = quote[f];
      return v === undefined ? "" : String(v);
    });
    invocation.setResult([header, [trimmedSymbol, ...values]]);
  };

  subscribeQuote(trimmedSymbol, listener);

  invocation.onCanceled = () => {
    canceled = true;
    unsubscribeQuote(trimmedSymbol, listener);
  };
}
/**
 * Server-free streaming test: counts up once per second.
 * @customfunction TESTSTREAM
 * @param start Starting value
 * @param invocation Custom function handler
 */
export function testStream(
  start: number,
  invocation: CustomFunctions.StreamingInvocation<number>
): void {
  let value = start;
  const timer = setInterval(() => {
    value += 1;
    invocation.setResult(value);
  }, 1000);

  invocation.onCanceled = () => clearInterval(timer);
}

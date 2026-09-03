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
 * Streams one live market-data field for a symbol.
 * Sample: =GETLIVEDATA("ACC.NS","LAST")
 * @customfunction GETLIVEDATA
 * @param {string} symbol Ticker symbol, e.g. ACC.NS
 * @param {string} fields Field name: LAST, OPEN, HIGH, LOW, or VOLUME
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Invocation
 */
export function getLiveData(
  symbol: string,
  fields: string,
  invocation: CustomFunctions.StreamingInvocation<string>
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

  if (requested.length !== 1) {
    invocation.setResult(
      new CustomFunctions.Error(
        CustomFunctions.ErrorCode.invalidValue,
        "Provide exactly one field: LAST, OPEN, HIGH, LOW, or VOLUME."
      )
    );
    return;
  }

  const field = requested[0];
  let canceled = false;

  invocation.setResult("Connecting...");

  const listener = (quote: Record<string, number | undefined>): void => {
    if (canceled) return;
    const value = quote[field];
    invocation.setResult(value === undefined ? "" : String(value));
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


/**
 * Get text values that spill down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillDown() {
  return [['first'], ['second'], ['third']];
}

/**
 * Get text values that spill to the right.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRight() {
  return [['first', 'second', 'third']];
}

/**
 * Get text values that spill both right and down.
 * @customfunction
 * @returns {string[][]} A dynamic array with multiple results.
 */
function spillRectangle() {
  return [
    ['apples', 1, 'pounds'],
    ['oranges', 3, 'pounds'],
    ['pears', 5, 'crates']
  ];
}


/**
 * Increment the cells with a given amount every second. Creates a dynamic spilled array with multiple results.
 * @customfunction
 * @param {number} amount The amount to add to the cell value on each increment.
 * @param {CustomFunctions.StreamingInvocation<number[][]>} invocation Parameter to send results to Excel or respond to the user canceling the function. A dynamic array.
 */
function increment(amount: number, invocation: CustomFunctions.StreamingInvocation<number[][]>): void {
  let firstResult = 0;
  let secondResult = 1;
  let thirdResult = 2;

  const timer = setInterval(() => {
    firstResult += amount;
    secondResult += amount;
    thirdResult += amount;
    invocation.setResult([[firstResult], [secondResult], [thirdResult]]);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}


/**
 * Calculate squares of input numbers.
 * @customfunction
 * @param {number[]} numbers Array of numbers to process.
 * @returns {any[][]} A dynamic array showing numbers and their squares.
 */
function calculateSquares(numbers: any[]) {
  // Create header row.
  const result = [['Number', 'Square']];

  // Process each number.
  numbers.forEach(row => {
    const num = Array.isArray(row) ? row[0] : row;
    result.push([num, num * num]);
  });

  return result;
}
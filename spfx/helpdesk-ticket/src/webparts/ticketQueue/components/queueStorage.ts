// The queue's remembered choices live in this browser only. Every access is
// guarded: a private window or blocked storage just means nothing is remembered.

export const FILTER_KEY = 'ticketQueue:filter:v1';
export const COLUMNS_KEY = 'ticketQueue:columns:v1';

export function readStored(key: string): string | undefined {
  try {
    const value = window.localStorage.getItem(key);
    return value === null ? undefined : value;
  } catch {
    return undefined;
  }
}

export function writeStored(key: string, value: string): void {
  try {
    window.localStorage.setItem(key, value);
  } catch {
    /* Storage is unavailable; the choice simply is not remembered. */
  }
}

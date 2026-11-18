/**
 * What happens to the device's copy of a till receipt when the receipt
 * changes.
 *
 * A receipt with lines is saved, so it survives a reload. One that has been
 * emptied by hand DROPS the stored copy: saving only non-empty receipts used
 * to leave the last line behind, and it came back on the next visit (and sat
 * in the unsaved-deliveries banner on /assets). A fresh receipt nobody has
 * scanned onto was never stored and needs nothing; and only THIS receipt's
 * copy is ever dropped.
 */
export const STORE = { SAVE: 'save', DROP: 'drop', NONE: 'none' };

export function storeAction(batch, storedId) {
  if (batch.drafts.length) return STORE.SAVE;
  return storedId && storedId === batch.id ? STORE.DROP : STORE.NONE;
}

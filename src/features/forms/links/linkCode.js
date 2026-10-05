/**
 * The short code at the end of a shared checklist link: `/c/7Kq2mX9aRt`.
 *
 * Ten characters of base62 is about 59 bits — short enough to paste into a
 * chat or read out, and far beyond what anybody could guess their way through
 * before the link expires. That is the whole of the protection, so the code
 * carries nothing else: no name, no row id, no date.
 */

const ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
export const LINK_CODE_LENGTH = 10;

// 248 is the largest multiple of 62 a byte can hold. A byte at or above it is
// thrown away rather than wrapped, or the first eight letters would turn up
// more often than the rest.
const LIMIT = 256 - (256 % ALPHABET.length);

const defaultRandom = (bytes) => globalThis.crypto.getRandomValues(bytes);

export function newLinkCode(random = defaultRandom) {
  let code = '';
  while (code.length < LINK_CODE_LENGTH) {
    const bytes = random(new Uint8Array(16));
    for (const byte of bytes) {
      if (byte >= LIMIT) continue;
      code += ALPHABET[byte % ALPHABET.length];
      if (code.length === LINK_CODE_LENGTH) break;
    }
  }
  return code;
}

/**
 * Checked before a code goes anywhere near a query. A code that is not ten
 * letters and digits cannot be one we made, and refusing it here means the
 * filter it would be pasted into never sees a quote.
 */
export function isLinkCode(value) {
  return typeof value === 'string' && /^[A-Za-z0-9]{10}$/.test(value);
}

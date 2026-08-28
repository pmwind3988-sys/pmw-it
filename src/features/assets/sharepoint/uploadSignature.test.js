import { describe, it, expect } from 'vitest';
import { dataUrlToBytes, signatureBlob } from './uploadSignature.js';

const png = 'data:image/png;base64,iVBORw0KGgo=';

describe('dataUrlToBytes', () => {
  it('gives back the bytes inside a data URL', () => {
    const bytes = dataUrlToBytes(png);

    expect(bytes).toBeInstanceOf(Uint8Array);
    // The PNG magic number, so this is the picture and not its base64.
    expect([...bytes.slice(0, 4)]).toEqual([0x89, 0x50, 0x4e, 0x47]);
  });

  it('is nothing at all for a skipped signature', () => {
    expect(dataUrlToBytes(null)).toBeNull();
    expect(dataUrlToBytes('')).toBeNull();
    expect(dataUrlToBytes('not-a-data-url')).toBeNull();
  });
});

describe('signatureBlob', () => {
  it('is a PNG ready to upload', () => {
    const blob = signatureBlob(png);

    expect(blob.type).toBe('image/png');
    expect(blob.size).toBeGreaterThan(0);
  });

  it('is nothing when nobody signed, so nothing is uploaded', () => {
    expect(signatureBlob(null)).toBeNull();
    expect(signatureBlob('data:image/png;base64,')).toBeNull();
  });
});

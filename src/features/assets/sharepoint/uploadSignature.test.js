import { describe, it, expect, vi } from 'vitest';
import { dataUrlToBytes, signatureBlob, uploadSignature } from './uploadSignature.js';

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

describe('uploadSignature', () => {
  /** One folder lookup and one upload, both answering as SharePoint would. */
  const okFetch = () => vi.fn(async (url) => {
    if (String(url).includes('RootFolder')) {
      return {
        ok: true,
        status: 200,
        json: async () => ({ d: { ServerRelativeUrl: '/sites/it/Photos' } }),
      };
    }
    return {
      ok: true,
      status: 200,
      json: async () => ({ d: { ServerRelativeUrl: '/sites/it/Photos/signature-amir-1.png' } }),
      text: async () => '',
    };
  });

  const upload = (extra) => uploadSignature({
    siteUrl: 'https://sp/sites/it',
    token: 't',
    digest: 'd',
    dataUrl: png,
    seed: 'amir',
    wait: async () => {},
    ...extra,
  });

  it('gives back where the signature landed', async () => {
    globalThis.fetch = okFetch();

    await expect(upload()).resolves.toBe('/sites/it/Photos/signature-amir-1.png');
  });

  it('tries again when the upload drops, rather than losing the signature', async () => {
    // The case that actually happens: store-room wifi drops one request while
    // the person is still standing at the desk. A signature only exists for
    // those few seconds, so one failure must not be the end of it.
    const good = okFetch();
    let calls = 0;
    globalThis.fetch = vi.fn(async (...args) => {
      calls += 1;
      if (calls === 1) throw new Error('network');
      return good(...args);
    });

    await expect(upload()).resolves.toBe('/sites/it/Photos/signature-amir-1.png');
  });

  it('gives up in the end, and says why', async () => {
    globalThis.fetch = vi.fn(async () => { throw new Error('network'); });

    await expect(upload({ attempts: 2 })).rejects.toThrow('network');
  });

  it('does not upload anything when nobody signed', async () => {
    globalThis.fetch = vi.fn();

    await expect(upload({ dataUrl: null })).resolves.toBe('');
    expect(globalThis.fetch).not.toHaveBeenCalled();
  });
});

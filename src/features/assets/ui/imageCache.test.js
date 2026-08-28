import {
  describe, it, expect, beforeEach, vi,
} from 'vitest';
import { loadImage, peekImage, clearImageCache } from './imageCache.js';

const KEY = 'https://sp/_api/photo-1.jpg';

beforeEach(() => {
  let next = 0;
  globalThis.URL.createObjectURL = vi.fn(() => {
    next += 1;
    return `blob:${next}`;
  });
  globalThis.URL.revokeObjectURL = vi.fn();
  clearImageCache();
});

describe('the picture cache', () => {
  it('fetches a picture once and answers the second ask from memory', async () => {
    const load = vi.fn(async () => new Blob(['bytes']));

    const first = await loadImage(KEY, load);
    const second = await loadImage(KEY, load);

    expect(second).toBe(first);
    // The whole point: opening an item twice must not download it twice.
    expect(load).toHaveBeenCalledTimes(1);
  });

  it('shares one request between two asks that overlap', async () => {
    const load = vi.fn(async () => new Blob(['bytes']));

    const [a, b] = await Promise.all([loadImage(KEY, load), loadImage(KEY, load)]);

    expect(a).toBe(b);
    expect(load).toHaveBeenCalledTimes(1);
  });

  it('has nothing to show before anything is fetched', () => {
    expect(peekImage(KEY)).toBeNull();
    expect(peekImage('')).toBeNull();
  });

  it('offers a fetched picture without a request', async () => {
    await loadImage(KEY, async () => new Blob(['bytes']));

    expect(peekImage(KEY)).toBe('blob:1');
  });

  it('lets a failed fetch be tried again rather than remembering the failure', async () => {
    const load = vi.fn()
      .mockRejectedValueOnce(new Error('500'))
      .mockResolvedValueOnce(new Blob(['bytes']));

    await expect(loadImage(KEY, load)).rejects.toThrow('500');
    await expect(loadImage(KEY, load)).resolves.toBe('blob:1');
  });

  it('revokes what it drops, so a long session does not leak a blob per photo', async () => {
    // Sixty-one pictures against a cache of sixty: the oldest has to go, and
    // going without being revoked is the leak this cache would otherwise be.
    for (let index = 0; index <= 60; index += 1) {
      await loadImage(`${KEY}?${index}`, async () => new Blob(['bytes']));
    }

    expect(globalThis.URL.revokeObjectURL).toHaveBeenCalledWith('blob:1');
    expect(peekImage(`${KEY}?0`)).toBeNull();
    expect(peekImage(`${KEY}?60`)).toBe('blob:61');
  });
});

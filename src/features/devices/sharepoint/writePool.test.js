import { describe, it, expect, vi } from 'vitest';
import { runPool, withRetry } from './writePool.js';

const response = (status, headers = {}) => ({
  ok: status >= 200 && status < 300,
  status,
  headers: { get: (name) => headers[name.toLowerCase()] ?? null },
});

describe('runPool', () => {
  it('returns a result per item, in input order', async () => {
    const results = await runPool([1, 2, 3], async (n) => n * 2, { concurrency: 2 });
    expect(results.map((r) => r.value)).toEqual([2, 4, 6]);
  });

  it('never runs more than `concurrency` workers at once', async () => {
    let active = 0;
    let peak = 0;
    await runPool([1, 2, 3, 4, 5, 6], async () => {
      active += 1;
      peak = Math.max(peak, active);
      await Promise.resolve();
      active -= 1;
    }, { concurrency: 2 });
    expect(peak).toBeLessThanOrEqual(2);
  });

  it('captures a failure without stopping the rest', async () => {
    const results = await runPool([1, 2, 3], async (n) => {
      if (n === 2) throw new Error('nope');
      return n;
    }, { concurrency: 3 });

    expect(results[1].error.message).toBe('nope');
    expect(results[0].value).toBe(1);
    expect(results[2].value).toBe(3);
  });

  it('reports progress as each item finishes', async () => {
    const seen = [];
    await runPool([1, 2, 3], async (n) => n, {
      concurrency: 1,
      onProgress: (done, total) => seen.push(`${done}/${total}`),
    });
    expect(seen).toEqual(['1/3', '2/3', '3/3']);
  });

  it('handles an empty batch without hanging', async () => {
    expect(await runPool([], async (n) => n, { concurrency: 4 })).toEqual([]);
  });
});

describe('withRetry', () => {
  it('returns a successful response without retrying', async () => {
    const attempt = vi.fn(async () => response(201));
    const result = await withRetry(attempt, { wait: async () => {} });
    expect(result.status).toBe(201);
    expect(attempt).toHaveBeenCalledTimes(1);
  });

  it('retries a 429 and honours Retry-After', async () => {
    const waits = [];
    const attempt = vi.fn()
      .mockResolvedValueOnce(response(429, { 'retry-after': '2' }))
      .mockResolvedValueOnce(response(201));

    const result = await withRetry(attempt, { wait: async (ms) => { waits.push(ms); } });

    expect(result.status).toBe(201);
    expect(waits).toEqual([2000]);
  });

  it('backs off exponentially when there is no Retry-After', async () => {
    const waits = [];
    const attempt = vi.fn(async () => response(503));

    await withRetry(attempt, { retries: 3, wait: async (ms) => { waits.push(ms); } });

    expect(attempt).toHaveBeenCalledTimes(3);
    expect(waits).toEqual([500, 1000]);
  });

  it('does not retry a 400 — a bad row will stay bad', async () => {
    const attempt = vi.fn(async () => response(400));
    const result = await withRetry(attempt, { wait: async () => {} });
    expect(result.status).toBe(400);
    expect(attempt).toHaveBeenCalledTimes(1);
  });
});

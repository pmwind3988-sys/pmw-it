const RETRYABLE = new Set([429, 503]);

const defaultWait = (ms) => new Promise((resolve) => { setTimeout(resolve, ms); });

export async function withRetry(attempt, { retries = 3, wait = defaultWait } = {}) {
  let response;

  for (let tryNumber = 1; tryNumber <= retries; tryNumber += 1) {
    response = await attempt();

    // A 400 means the row is wrong, not that SharePoint is busy. Retrying it
    // just costs the user three times as long to see the same error.
    if (response.ok || !RETRYABLE.has(response.status)) return response;
    if (tryNumber === retries) return response;

    const retryAfter = Number(response.headers?.get?.('Retry-After'));
    await wait(Number.isFinite(retryAfter) && retryAfter > 0
      ? retryAfter * 1000
      : 500 * 2 ** (tryNumber - 1));
  }

  return response;
}

/**
 * Four writes in flight rather than SharePoint's multipart $batch: a hand-built
 * multipart body is easy to get subtly wrong, and this imports 200 machines in
 * well under a minute. $batch stays available as a later optimisation.
 */
export async function runPool(items, worker, { concurrency = 4, onProgress } = {}) {
  const results = new Array(items.length);
  let next = 0;
  let done = 0;

  const runner = async () => {
    while (next < items.length) {
      const index = next;
      next += 1;

      try {
        results[index] = { item: items[index], value: await worker(items[index]), error: null };
      } catch (error) {
        results[index] = { item: items[index], value: null, error };
      }

      done += 1;
      onProgress?.(done, items.length);
    }
  };

  await Promise.all(
    Array.from({ length: Math.min(concurrency, items.length) }, () => runner()),
  );

  return results;
}

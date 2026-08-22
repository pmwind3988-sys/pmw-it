import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

// Reads embed.js as TEXT rather than importing it: importing would pull
// in the model runtime, which is the one thing the test suite must never
// do. What is being checked is a promise about configuration, and the
// configuration is legible without running it.
const read = (relative) => readFileSync(
  fileURLToPath(new URL(relative, import.meta.url)),
  'utf8',
// Normalised, because git checks these files out with CRLF on Windows
// and every pattern below is written with plain newlines.
).replace(/\r\n/g, '\n');

const SOURCE = read('./embed.js');
const GITIGNORE = read('../../../../.gitignore');
const FETCH_SCRIPT = read('../../../../scripts/fetch-model.mjs');

describe('the model stays on this origin', () => {
  it('refuses remote models and allows local ones', () => {
    // Both lines are needed. `allowLocalModels` defaults to FALSE in the
    // browser, so switching remote off alone disables both.
    expect(SOURCE).toMatch(/env\.allowRemoteModels\s*=\s*false/);
    expect(SOURCE).toMatch(/env\.allowLocalModels\s*=\s*true/);
  });

  it('never names a host', () => {
    const hosts = SOURCE.match(/https?:\/\/[^\s'"`]+/g) ?? [];
    expect(hosts).toEqual([]);
  });
});

describe('every path it points at is one the build produces', () => {
  // The bug this pins: `wasmPaths` was set to '/ort/…', a folder that
  // only ever existed because it had been copied there by hand. It is
  // gitignored and the fetch script does not create it, so a fresh
  // checkout served index.html for the .wasm and the runtime aborted
  // with "expected magic word 00 61 73 6d, found 3c 21 64 6f" -- which
  // is "<!do", the SPA fallback answering for a missing file.
  const publicPaths = [...SOURCE.matchAll(/'(\/[a-z0-9/_.-]+)'/gi)]
    .map((m) => m[1])
    .filter((p) => !p.startsWith('//'));

  it('only points at /models/, which the fetch script fills', () => {
    expect(publicPaths).toEqual(['/models/']);
    expect(FETCH_SCRIPT).toContain("join(ROOT, 'public', 'models'");
  });

  it('does not depend on a gitignored folder nothing creates', () => {
    const ignored = GITIGNORE.split('\n')
      .map((l) => l.trim())
      .filter((l) => l.startsWith('public/'))
      .map((l) => `/${l.slice('public/'.length)}`);

    for (const path of publicPaths) {
      const underIgnored = ignored.find((i) => path.startsWith(i));
      if (!underIgnored) continue;
      // Ignored is fine only if the build script recreates it.
      expect(FETCH_SCRIPT).toContain('public');
      expect(underIgnored).toBe('/models/');
    }
  });
});

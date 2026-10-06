import { describe, it, expect, afterEach, vi } from 'vitest';
import { toStandardItem, fromStandardItem } from './standardsSchema.js';
import { pickStandard, readStandards, hasPermission, canEditStandards } from './readStandards.js';
import { saveStandard } from './saveStandard.js';
import { defaultStandard, DEFAULT_STANDARD } from '../standards/defaultStandard.js';

const reply = (body, status = 200) => ({ ok: status < 300, status, json: async () => body, text: async () => 'x', headers: { get: () => null } });
const row = (version, standard) => ({ Id: version, ...toStandardItem({ version, standard, savedBy: 'aisyah', summary: `v${version}`, savedOn: Date.UTC(2026, 9, version) }) });

describe('standards rows', () => {
  it('round-trip a standard through a list item', () => {
    const back = fromStandardItem(row(4, defaultStandard()));
    expect(back).toMatchObject({ id: 4, version: 4, savedBy: 'aisyah', summary: 'v4', savedOn: Date.UTC(2026, 9, 4) });
    expect(back.standard).toEqual(defaultStandard());
  });

  it('read an unparsable row as no standard', () => {
    expect(fromStandardItem({ Id: 2, Title: 'v2', Standard: '{nope' }).standard).toBeNull();
  });
});

describe('pickStandard', () => {
  it('uses the newest valid version', () => {
    const picked = pickStandard([row(5, defaultStandard()), row(4, defaultStandard())].map(fromStandardItem));
    expect(picked.version).toBe(5);
    expect(picked.note).toBeNull();
    expect(picked.history).toHaveLength(2);
  });

  it('skips a broken newest version and says so', () => {
    const bad = { ...defaultStandard(), schema: 9 };
    const picked = pickStandard([row(6, bad), row(5, defaultStandard())].map(fromStandardItem));
    expect(picked.version).toBe(5);
    expect(picked.note).toBe('Version 6 could not be used, so version 5 is in force.');
  });

  it('names the row in force, which is not always the newest', () => {
    const bad = { ...defaultStandard(), schema: 9 };
    expect(pickStandard([row(6, bad), row(5, defaultStandard())].map(fromStandardItem)).inForceId).toBe(5);
    expect(pickStandard([]).inForceId).toBeNull();
  });

  it('falls back to the default standard with nothing saved', () => {
    expect(pickStandard([])).toMatchObject({ standard: DEFAULT_STANDARD, version: 0, note: null });
  });
});

describe('reading and permissions', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('reads the newest twenty, newest first', async () => {
    let asked = '';
    vi.stubGlobal('fetch', async (url) => { asked = decodeURIComponent(url); return reply({ d: { results: [row(2, defaultStandard())] } }); });
    const result = await readStandards('https://x', 't');
    expect(asked).toContain('$orderby=Id desc&$top=20');
    expect(result.version).toBe(2);
  });

  it('treats a list that does not exist yet as the default standard', async () => {
    vi.stubGlobal('fetch', async () => reply({}, 404));
    expect((await readStandards('https://x', 't')).version).toBe(0);
  });

  it('reads a permission bit', () => {
    expect(hasPermission({ Low: '2', High: '0' }, 2)).toBe(true);
    expect(hasPermission({ Low: '1', High: '0' }, 2)).toBe(false);
    expect(hasPermission({ Low: String(0x800), High: '0' }, 12)).toBe(true);
  });

  it('may edit when the list grants AddListItems', async () => {
    vi.stubGlobal('fetch', async () => reply({ d: { EffectiveBasePermissions: { Low: '3', High: '0' } } }));
    expect(await canEditStandards('https://x', 't')).toBe(true);
  });

  it('falls back to "may I create lists" before the list exists', async () => {
    vi.stubGlobal('fetch', async (url) => (url.includes("getByTitle")
      ? reply({}, 404)
      : reply({ d: { EffectiveBasePermissions: { Low: '0', High: '0' } } })));
    expect(await canEditStandards('https://x', 't')).toBe(false);
  });
});

describe('saveStandard', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('refuses an invalid standard before touching SharePoint', async () => {
    const fetch = vi.fn();
    vi.stubGlobal('fetch', fetch);
    const bad = defaultStandard();
    bad.profiles.heavy.ram.criticalBelow = 99;
    await expect(saveStandard({ siteUrl: 'https://x', token: 't', standard: bad, version: 3, summary: 's', savedBy: 'me' }))
      .rejects.toThrow(/in order/);
    expect(fetch).not.toHaveBeenCalled();
  });

  it('turns a 403 into a sentence about edit rights', async () => {
    vi.stubGlobal('fetch', async (url, init = {}) => {
      if (url.endsWith('/_api/contextinfo')) return reply({ d: { GetContextWebInformation: { FormDigestValue: 'D' } } });
      if (url.includes('/fields?')) return reply({ d: { results: [] } });
      if (init.method === 'POST' && url.includes('/items')) return reply({}, 403);
      return reply({ d: {} });
    });
    await expect(saveStandard({ siteUrl: 'https://x', token: 't', standard: defaultStandard(), version: 3, summary: 's', savedBy: 'me' }))
      .rejects.toThrow('You do not have edit rights on the IT Device Standards list. Ask IT.');
  });
});

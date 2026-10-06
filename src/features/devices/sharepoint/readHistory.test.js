import { describe, it, expect, afterEach, vi } from 'vitest';
import { readChanges } from './readHistory.js';

const SITE = 'https://contoso.sharepoint.com/sites/it';
const reply = (body, status = 200) => ({ ok: status < 300, status, json: async () => body, text: async () => '' });

describe('readChanges', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('reads the whole list and matches locally when DeviceId does not exist yet', async () => {
    const urls = [];
    vi.stubGlobal('fetch', async (url) => {
      urls.push(decodeURIComponent(url));
      if (urls.length === 1) return reply({}, 400);
      return reply({ d: { results: [
        { Title: 'AMIR-HP', FieldName: 'owner', NewValue: 'A' },
        { Title: 'OTHER-PC', FieldName: 'owner', NewValue: 'B' },
      ] } });
    });
    const rows = await readChanges(SITE, 't', { id: 4, computerName: 'AMIR-HP' });
    expect(urls).toHaveLength(2);
    expect(urls[0]).toContain('DeviceId eq 4');
    expect(urls[1]).not.toContain('$filter');
    expect(rows.map((r) => r.newValue)).toEqual(['A']);
  });

  it('survives the 5,000-item threshold, which SharePoint answers with a 500', async () => {
    const urls = [];
    vi.stubGlobal('fetch', async (url) => {
      urls.push(decodeURIComponent(url));
      if (urls.length === 1) return reply({}, 500);
      if (urls.length === 2) return reply({ d: { results: [{ Title: 'X', DeviceId: 4, NewValue: 'p1' }], __next: `${SITE}/next` } });
      return reply({ d: { results: [{ Title: 'Y', DeviceId: 5, NewValue: 'other' }, { Title: 'X', DeviceId: 4, NewValue: 'p2' }] } });
    });
    const rows = await readChanges(SITE, 't', { id: 4, computerName: 'X' });
    expect(urls).toHaveLength(3);
    expect(urls[1]).toContain('$top=5000');
    expect(rows.map((r) => r.newValue)).toEqual(['p1', 'p2']);
  });

  it('still reports a failure when the unfiltered read is refused too', async () => {
    vi.stubGlobal('fetch', async () => reply({}, 500));
    await expect(readChanges(SITE, 't', { id: 4, computerName: 'X' })).rejects.toThrow('(500)');
  });

  it("keeps another machine's rows out when a name is reused", async () => {
    vi.stubGlobal('fetch', async () => reply({ d: { results: [
      { Title: 'AMIR-HP', DeviceId: 9, FieldName: 'owner', NewValue: 'old machine' },
      { Title: 'AMIR-HP', DeviceId: 4, FieldName: 'owner', NewValue: 'mine' },
      { Title: 'AMIR-HP', DeviceId: null, FieldName: 'owner', NewValue: 'legacy' },
      { Title: 'RENAMED', DeviceId: 4, FieldName: 'status', NewValue: 'renamed' },
    ] } }));
    const rows = await readChanges(SITE, 't', { id: 4, computerName: 'AMIR-HP' });
    expect(rows.map((r) => r.newValue)).toEqual(['mine', 'legacy', 'renamed']);
    expect(rows[0].deviceId).toBe(4);
  });
});

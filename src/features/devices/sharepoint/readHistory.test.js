import { describe, it, expect, afterEach, vi } from 'vitest';
import { readChanges } from './readHistory.js';

const SITE = 'https://contoso.sharepoint.com/sites/it';
const reply = (body, status = 200) => ({ ok: status < 300, status, json: async () => body, text: async () => '' });

describe('readChanges', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('retries with the Title-only filter when DeviceId does not exist yet', async () => {
    const urls = [];
    vi.stubGlobal('fetch', async (url) => {
      urls.push(decodeURIComponent(url));
      if (urls.length === 1) return reply({}, 400);
      return reply({ d: { results: [{ Title: 'AMIR-HP', FieldName: 'owner', NewValue: 'A' }] } });
    });
    const rows = await readChanges(SITE, 't', { id: 4, computerName: 'AMIR-HP' });
    expect(urls).toHaveLength(2);
    expect(urls[0]).toContain('DeviceId eq 4');
    expect(urls[1]).not.toContain('DeviceId');
    expect(rows).toHaveLength(1);
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

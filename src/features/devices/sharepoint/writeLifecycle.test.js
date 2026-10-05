import { describe, it, expect, afterEach, vi } from 'vitest';
import { lifecycleItem, stintWrites, performLifecycle } from './writeLifecycle.js';

const SITE = 'https://contoso.sharepoint.com/sites/it';
const NOW = Date.UTC(2026, 9, 5, 4);

describe('lifecycleItem', () => {
  it('maps only the fields the action changed', () => {
    expect(lifecycleItem({
      owner: null, department: null, status: 'Spare', statusChangedOn: NOW, manualFields: ['deviceType'],
    })).toEqual({
      Owner: '', Department: '', Status: 'Spare',
      StatusChangedOn: '2026-10-05T04:00:00.000Z', ManualFields: 'deviceType',
    });
  });
});

describe('stintWrites', () => {
  it('merges the end onto a stint that exists', () => {
    const writes = stintWrites({ close: { stint: { id: 30 }, endedOn: NOW, endReason: 'To stash', note: null }, open: null });
    expect(writes).toEqual([{ id: 30, body: { EndedOn: '2026-10-05T04:00:00.000Z', EndReason: 'To stash' } }]);
  });

  it('writes a legacy stint whole, already ended', () => {
    const [write] = stintWrites({
      close: { stint: { id: null, deviceId: 4, owner: 'Amir', assignedOn: NOW - 1000, assignedOnApprox: true }, endedOn: NOW, endReason: 'Retired', note: 'old' },
      open: null,
    });
    expect(write.id).toBeNull();
    expect(write.body).toMatchObject({ DeviceId: 4, Owner: 'Amir', AssignedOnApprox: true, EndReason: 'Retired', Note: 'old' });
  });
});

function fakeSharePoint({ device, stints = [], failStints = false } = {}) {
  const calls = [];
  const reply = (body, status = 200) => ({
    ok: status < 300, status, json: async () => body, text: async () => 'boom', headers: { get: () => null },
  });
  return {
    calls,
    fetch: async (url, init = {}) => {
      const method = init.method ?? 'GET';
      calls.push({ url: decodeURIComponent(url), method, headers: init.headers ?? {}, body: init.body ? JSON.parse(init.body) : undefined });
      if (url.endsWith('/_api/contextinfo')) return reply({ d: { GetContextWebInformation: { FormDigestValue: 'D' } } });
      if (method === 'GET' && url.includes('IT%20Device%20List')) return reply({ d: device });
      if (method === 'GET' && url.includes('IT%20Device%20Assignments')) return reply({ d: { results: stints } });
      if (failStints && url.includes('IT%20Device%20Assignments')) return reply({}, 500);
      return reply({ Id: 99 }, 201);
    },
  };
}

const posts = (calls) => calls.filter((c) => c.method === 'POST' && !c.url.endsWith('/contextinfo'));

describe('performLifecycle', () => {
  afterEach(() => vi.unstubAllGlobals());

  const device = { Id: 4, Title: 'AMIR-HP', Owner: 'Amir', Location: 'F1', Department: 'ENGINEERING', Created: '2026-08-21T00:00:00Z' };

  it('writes the device row, then the stints, then the change log', async () => {
    const sp = fakeSharePoint({ device, stints: [{ Id: 30, DeviceId: 4, Owner: 'Amir', AssignedOn: '2026-02-12T00:00:00Z' }] });
    vi.stubGlobal('fetch', sp.fetch);

    await performLifecycle({
      siteUrl: SITE, token: 't', deviceId: 4, action: 'changeOwner',
      input: { owner: 'Aisyah', location: 'F1', department: 'ENGINEERING' }, expectedStatus: 'In use', recordedBy: 'it', now: NOW,
    });

    const order = posts(sp.calls).map((c) => {
      if (c.url.includes('IT Device List')) return 'device';
      if (c.url.includes('IT Device Assignments')) return 'stint';
      return 'change';
    });
    expect(order).toEqual(['device', 'stint', 'stint', 'change']);
    expect(posts(sp.calls)[1].url).toContain('items(30)');
    expect(posts(sp.calls)[3].body).toMatchObject({ DeviceId: 4, FieldName: 'owner', NewValue: 'Aisyah' });
  });

  it('refuses and writes nothing when the screen was stale', async () => {
    const sp = fakeSharePoint({ device: { ...device, Status: 'Spare', Owner: '' } });
    vi.stubGlobal('fetch', sp.fetch);

    await expect(performLifecycle({
      siteUrl: SITE, token: 't', deviceId: 4, action: 'toStash', expectedStatus: 'In use', now: NOW,
    })).rejects.toThrow(/someone else/);
    expect(posts(sp.calls)).toHaveLength(0);
  });

  it('reports stints it could not write instead of failing the whole action', async () => {
    const sp = fakeSharePoint({ device, failStints: true });
    vi.stubGlobal('fetch', sp.fetch);

    const result = await performLifecycle({
      siteUrl: SITE, token: 't', deviceId: 4, action: 'retire', expectedStatus: 'In use', now: NOW,
    });
    expect(result.pendingStints).toHaveLength(1);
  });
});

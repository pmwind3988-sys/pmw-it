import { describe, it, expect, afterEach, vi } from 'vitest';
import { lifecycleItem, stintWrites, performLifecycle } from './writeLifecycle.js';

vi.mock('./provisionLists.js', () => ({ provisionLists: vi.fn(async () => 'D') }));

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

function fakeSharePoint({
  device, stints = [], failStints = false, failStintAt = null, stintMode = 'status', failDevice = false, failLog = false,
} = {}) {
  const calls = [];
  let stintPosts = 0;
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
      if (failDevice && url.includes('IT%20Device%20List')) return reply({}, 400);
      if (failLog && url.includes('IT%20Device%20Changes')) return reply({}, 400);
      if (failStintAt !== null && url.includes('IT%20Device%20Assignments')) {
        const mine = stintPosts;
        stintPosts += 1;
        if (mine === failStintAt) {
          if (stintMode === 'throw') throw new TypeError('Failed to fetch');
          return reply({}, 400);
        }
      }
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

  const move = { siteUrl: SITE, token: 't', deviceId: 4, action: 'changeOwner', input: { owner: 'Aisyah', location: 'F1', department: 'ENGINEERING' }, expectedStatus: 'In use', now: NOW };
  const legacy = [{ Id: 30, DeviceId: 4, Owner: 'Amir', AssignedOn: '2026-02-12T00:00:00Z' }];

  it('writes no stints and no log when the device row fails', async () => {
    const sp = fakeSharePoint({ device, stints: legacy, failDevice: true });
    vi.stubGlobal('fetch', sp.fetch);
    await expect(performLifecycle(move)).rejects.toThrow(/Could not save the change/);
    expect(posts(sp.calls).filter((c) => c.url.includes('Assignments') || c.url.includes('Changes'))).toHaveLength(0);
  });

  it.each(['status', 'throw'])('leaves exactly the unwritten stint pending when the second write fails (%s)', async (mode) => {
    const sp = fakeSharePoint({ device, stints: legacy, failStintAt: 1, stintMode: mode });
    vi.stubGlobal('fetch', sp.fetch);
    const result = await performLifecycle(move);
    expect(result.pendingStints).toHaveLength(1);
    expect(result.pendingStints[0].id).toBeNull();
  });

  it('returns logFailed instead of throwing when only the change log fails', async () => {
    const sp = fakeSharePoint({ device, stints: legacy, failLog: true });
    vi.stubGlobal('fetch', sp.fetch);
    const result = await performLifecycle(move);
    expect(result.logFailed).toBe(true);
    expect(result.pendingStints).toHaveLength(0);
  });

  it('provisions once per page load, on the first call only', async () => {
    vi.resetModules();
    const { provisionLists } = await import('./provisionLists.js');
    provisionLists.mockClear();
    const { performLifecycle: fresh } = await import('./writeLifecycle.js');
    const sp = fakeSharePoint({ device, stints: legacy });
    vi.stubGlobal('fetch', sp.fetch);
    await fresh(move);
    const digestCallsAfterFirst = sp.calls.filter((c) => c.url.endsWith('/_api/contextinfo')).length;
    await fresh(move);
    expect(provisionLists).toHaveBeenCalledTimes(1);
    // The second action fetches its own digest rather than reusing a cached one.
    expect(sp.calls.filter((c) => c.url.endsWith('/_api/contextinfo')).length).toBe(digestCallsAfterFirst + 1);
  });

  it('forgets a failed provisioning so the next press retries', async () => {
    vi.resetModules();
    const { provisionLists } = await import('./provisionLists.js');
    provisionLists.mockClear();
    provisionLists.mockRejectedValueOnce(new Error('offline'));
    const { performLifecycle: fresh } = await import('./writeLifecycle.js');
    const sp = fakeSharePoint({ device, stints: legacy });
    vi.stubGlobal('fetch', sp.fetch);
    await expect(fresh(move)).rejects.toThrow('offline');
    await fresh(move);
    expect(provisionLists).toHaveBeenCalledTimes(2);
  });
});

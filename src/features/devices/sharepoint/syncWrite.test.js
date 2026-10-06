import { describe, it, expect, afterEach, vi } from 'vitest';

vi.mock('./provisionLists.js', () => ({ provisionLists: vi.fn(async () => 'D') }));
vi.mock('./readDevices.js', () => ({ readAllDevices: vi.fn() }));
vi.mock('./readAssignments.js', () => ({ readAssignments: vi.fn() }));

const { syncDevices } = await import('./syncDevices.js');
const { readAllDevices } = await import('./readDevices.js');
const { readAssignments } = await import('./readAssignments.js');

const SITE = 'https://contoso.sharepoint.com/sites/it';
const scan = (over) => ({
  computerName: 'PC1', owner: 'Ali', department: 'SALES', deviceType: 'Laptop', serialNumber: 'S1',
  scannedOn: Date.UTC(2026, 9, 6), sourceFileName: 'PC1.txt', ...over,
});

function stubFetch({ failOn = () => false } = {}) {
  const calls = [];
  let nextId = 55;
  vi.stubGlobal('fetch', async (url, init = {}) => {
    const decoded = decodeURIComponent(url);
    const body = init.body ? JSON.parse(init.body) : undefined;
    calls.push({ url: decoded, method: init.method ?? 'GET', headers: init.headers ?? {}, body });
    const fail = failOn(decoded, init);
    if (fail) return { ok: false, status: 400, json: async () => ({}), text: async () => 'no', headers: { get: () => null } };
    const created = decoded.endsWith('/items') ? { Id: nextId++ } : {};
    return { ok: true, status: 201, json: async () => created, text: async () => '', headers: { get: () => null } };
  });
  return calls;
}

const inList = (c, name) => c.url.includes(name);

describe('syncDevices write path', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('gives the first stint of a new machine the Id SharePoint assigned', async () => {
    readAllDevices.mockResolvedValue([]);
    readAssignments.mockResolvedValue([]);
    const calls = stubFetch();
    await syncDevices({ siteUrl: SITE, token: 't', devices: [scan()], changedBy: 'it' });
    const stint = calls.find((c) => inList(c, 'IT Device Assignments'));
    expect(stint.body).toMatchObject({ DeviceId: 55, Owner: 'Ali' });
  });

  it('writes no stint for a device write that failed, even when a same-named insert succeeded', async () => {
    const old = {
      id: 8, computerName: 'PC-OLD', serialNumber: 'S0', owner: 'Ali', department: 'SALES', deviceType: 'Laptop', createdOn: 1,
    };
    readAllDevices.mockResolvedValue([old]);
    readAssignments.mockResolvedValue([{
      id: 30, deviceId: 8, owner: 'Ali', location: null, department: 'SALES', assignedOn: 1000, endedOn: null,
    }]);
    // The retirement of row 8 fails; the same-named new machine goes in fine.
    const calls = stubFetch({ failOn: (url, init) => init.method === 'POST' && url.includes('items(8)') });
    const result = await syncDevices({
      siteUrl: SITE, token: 't', changedBy: 'it',
      devices: [scan({ computerName: 'PC-OLD', serialNumber: 'S2', sourceFileName: 'new.txt' })],
      answers: { 'new.txt|8': 'stash' },
    });
    expect(result.results.find((r) => r.action === 'retire').error).toBeTruthy();
    const stintCalls = calls.filter((c) => inList(c, 'IT Device Assignments'));
    expect(stintCalls.some((c) => c.url.includes('items(30)'))).toBe(false);
    expect(stintCalls.some((c) => c.body?.DeviceId === 55)).toBe(true);
  });
});

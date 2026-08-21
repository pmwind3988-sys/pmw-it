import {
  describe, it, expect, afterEach, vi,
} from 'vitest';
import { syncDevices } from './syncDevices.js';
import { DEVICE_COLUMNS, CHANGE_COLUMNS } from './deviceSchema.js';

const SITE = 'https://contoso.sharepoint.com/sites/it';

const device = (overrides) => ({
  computerName: 'PC1', owner: 'Ali', department: 'SALES', deviceType: 'Laptop',
  computerModel: 'HP 15', windowsVersion: 'Microsoft Windows 11 Pro', osSupported: true,
  cpuModel: 'i5', cpuAgeBand: 'Current', installedRamGB: 8, ramType: 'DDR4', ramSlotsUsed: 2,
  storageTotalGB: 477, storageType: 'SSD only', antivirusStatus: 'Active', riskLevel: 'Watch',
  scannedOn: Date.UTC(2026, 7, 19, 1, 18), sourceFileName: 'PC1_.txt',
  ...overrides,
});

/**
 * Every column already exists and is correctly named, so provisioning is pure
 * checking. That is the second-run shape, and the one where a bar driven only
 * by row writes would sit still through the slowest part.
 */
function fakeSharePoint({ items = [] } = {}) {
  const reply = (body, status = 200) => ({
    ok: status >= 200 && status < 300,
    status,
    json: async () => body,
    text: async () => JSON.stringify(body),
    headers: { get: () => null },
  });

  const fields = [...DEVICE_COLUMNS, ...CHANGE_COLUMNS].map((column) => ({
    InternalName: column.StaticName, Title: column.Title,
  }));

  const written = [];

  return {
    written,
    fetch: async (url, init = {}) => {
      if (url.endsWith('/_api/contextinfo')) {
        return reply({ d: { GetContextWebInformation: { FormDigestValue: 'DIGEST' } } });
      }
      if (url.includes('/fields?$select=')) return reply({ d: { results: fields } });
      if (url.includes('/items?')) return reply({ d: { results: items } });
      if (url.includes('/items')) {
        written.push({ url, body: JSON.parse(init.body) });
        return reply({}, 201);
      }
      return reply({});
    },
  };
}

const phasesFrom = (updates) => [...new Set(updates.map((u) => u.phase))];

describe('syncDevices progress phases', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('reports provisioning and reading before any row is written', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    const updates = [];
    await syncDevices({
      siteUrl: SITE,
      token: 'token',
      devices: [device()],
      changedBy: 'me@pmw-group.com',
      onProgress: (update) => updates.push(update),
    });

    // The order matters: a first save spends most of its time in the first two.
    expect(phasesFrom(updates)).toEqual(['provisioning', 'reading', 'writing']);
  });

  it('counts columns during provisioning, not rows', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    const updates = [];
    await syncDevices({
      siteUrl: SITE, token: 'token', devices: [device()], onProgress: (u) => updates.push(u),
    });

    const provisioning = updates.filter((u) => u.phase === 'provisioning');
    const total = DEVICE_COLUMNS.length + CHANGE_COLUMNS.length;
    expect(provisioning.at(-1)).toEqual({ phase: 'provisioning', done: total, total });
  });

  it('reports the row writes with the row count', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    const updates = [];
    await syncDevices({
      siteUrl: SITE,
      token: 'token',
      devices: [device(), device({ computerName: 'PC2', sourceFileName: 'PC2_.txt' })],
      onProgress: (u) => updates.push(u),
    });

    expect(updates.filter((u) => u.phase === 'writing').at(-1))
      .toEqual({ phase: 'writing', done: 2, total: 2 });
  });

  it('reports a logging phase only when there are changes to log', async () => {
    const existing = {
      Id: 7, Title: 'PC1', Owner: 'Ali', Department: 'SALES', DeviceType: 'Laptop',
      ComputerModel: 'HP 15', WindowsVersion: 'Microsoft Windows 11 Pro', OsSupported: true,
      CpuModel: 'i5', CpuAgeBand: 'Current', InstalledRamGB: 8, RamType: 'DDR4',
      RamSlotsUsed: 2, StorageTotalGB: 477, StorageType: 'SSD only',
      AntivirusStatus: 'Active', RiskLevel: 'Watch',
    };

    const unchanged = fakeSharePoint({ items: [existing] });
    vi.stubGlobal('fetch', unchanged.fetch);
    const quiet = [];
    await syncDevices({
      siteUrl: SITE, token: 'token', devices: [device()], onProgress: (u) => quiet.push(u),
    });
    expect(phasesFrom(quiet)).not.toContain('logging');
    expect(unchanged.written).toHaveLength(0);

    const upgraded = fakeSharePoint({ items: [existing] });
    vi.stubGlobal('fetch', upgraded.fetch);
    const loud = [];
    await syncDevices({
      siteUrl: SITE,
      token: 'token',
      devices: [device({ installedRamGB: 16 })],
      onProgress: (u) => loud.push(u),
    });
    expect(phasesFrom(loud)).toContain('logging');
  });

  it('runs without a progress callback', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    const outcome = await syncDevices({ siteUrl: SITE, token: 'token', devices: [device()] });
    expect(outcome.results).toEqual([{ computerName: 'PC1', action: 'insert', error: null }]);
  });
});

import {
  describe, it, expect, afterEach, vi,
} from 'vitest';
import { updateAsset, isMissingColumn, planEdit } from './updateAsset.js';
import { serialiseUnits, parseUnits } from '../units.js';

const SITE = 'https://contoso.sharepoint.com/sites/it';

const bulkRow = (overrides) => ({
  id: 12,
  title: 'Lenovo Tab M11',
  assetKey: 'TAB-M11',
  category: 'Tablet',
  trackingMode: 'Bulk',
  quantity: 2,
  manualFields: [],
  ...overrides,
});

/** Exactly what the tenant answered when `Units` had never been created. */
const MISSING_UNITS = JSON.stringify({
  'odata.error': {
    code: '-1, Microsoft.SharePoint.Client.InvalidClientQueryException',
    message: {
      lang: 'en-US',
      value: "The property 'Units' does not exist on type "
        + "'SP.Data.IT_x0020_Asset_x0020_RegisterListItem'. Make sure to only use property "
        + 'names that are defined by the type.',
    },
  },
});

describe('isMissingColumn', () => {
  it('recognises the tenant complaining about a column it has never had', () => {
    expect(isMissingColumn(400, MISSING_UNITS)).toBe(true);
  });

  it('does not mistake an ordinary refusal for one', () => {
    expect(isMissingColumn(400, 'The request is invalid')).toBe(false);
    expect(isMissingColumn(403, MISSING_UNITS)).toBe(false);
    expect(isMissingColumn(500, '')).toBe(false);
  });
});

/**
 * Only what these tests actually ask of SharePoint. `provisionAssets` is the
 * subject here — that it RUNS, and that the write is tried again after it —
 * not what it creates; `provision.test.js` already covers the creating.
 */
function fakeSharePoint({ missingUntilProvisioned = false } = {}) {
  const calls = [];
  let provisioned = !missingUntilProvisioned;

  const reply = (status, body = '') => ({
    ok: status >= 200 && status < 300,
    status,
    json: async () => ({ d: { results: [] } }),
    text: async () => body,
    headers: { get: () => null },
  });

  return {
    calls,
    get provisioned() { return provisioned; },
    fetch: async (url, init = {}) => {
      calls.push({ url, method: init.method });

      if (url.endsWith('/_api/contextinfo')) {
        return {
          ...reply(200),
          json: async () => ({ d: { GetContextWebInformation: { FormDigestValue: 'D' } } }),
        };
      }

      // Provisioning walks the fields of each list. Answering it at all is
      // what marks the schema as brought up to date.
      if (url.includes('/fields')) {
        provisioned = true;
        return reply(200);
      }
      if (url.includes('/views') || url.includes('/RootFolder')) return reply(200);
      if (url.includes("lists/getByTitle") && init.method === undefined) return reply(200);

      if (url.includes('items(12)') && !provisioned) return reply(400, MISSING_UNITS);

      return reply(200);
    },
  };
}

const rowWrites = (calls) => calls.filter((call) => call.url.includes('items(12)'));

describe('updateAsset', () => {
  afterEach(() => { vi.unstubAllGlobals(); });

  /**
   * The list outlives every release, so a column added to the app's schema
   * afterwards exists in the code and not in the tenant. Every save of a bulk
   * row failed on that, with a message about `SP.Data...ListItem` that told
   * the person holding the phone nothing they could act on.
   */
  it('creates the missing column and saves, rather than refusing', async () => {
    const sp = fakeSharePoint({ missingUntilProvisioned: true });
    vi.stubGlobal('fetch', sp.fetch);

    const result = await updateAsset({
      siteUrl: SITE,
      token: 't',
      existing: bulkRow(),
      edits: { units: serialiseUnits([{ index: 0, serialNumber: 'HA2KJDSW' }]) },
      changedBy: 'me@pmw-group.com',
    });

    expect(result.repaired).toBe(true);
    // Tried, repaired, tried again — and the second one landed.
    expect(rowWrites(sp.calls)).toHaveLength(2);
  });

  it('does not provision when nothing is wrong', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    const result = await updateAsset({
      siteUrl: SITE,
      token: 't',
      existing: bulkRow(),
      edits: { units: serialiseUnits([{ index: 0, serialNumber: 'HA2KJDSW' }]) },
      changedBy: 'me@pmw-group.com',
    });

    expect(result.repaired).toBe(false);
    expect(rowWrites(sp.calls)).toHaveLength(1);
    expect(sp.calls.some((call) => call.url.includes('/fields'))).toBe(false);
  });

  it('still gives up on a refusal it cannot repair', async () => {
    vi.stubGlobal('fetch', async (url) => {
      if (url.endsWith('/_api/contextinfo')) {
        return {
          ok: true,
          status: 200,
          json: async () => ({ d: { GetContextWebInformation: { FormDigestValue: 'D' } } }),
          headers: { get: () => null },
        };
      }
      return {
        ok: false, status: 403, text: async () => 'Access denied', headers: { get: () => null },
      };
    });

    await expect(updateAsset({
      siteUrl: SITE,
      token: 't',
      existing: bulkRow(),
      edits: { location: 'Store room' },
    })).rejects.toThrow(/403/);
  });

  /**
   * A photograph taken of item 3 changes the row and produces no change-log
   * line, because the log deliberately ignores photos. Without the unit
   * records being compared in their own right, the save would decide nothing
   * had happened and quietly drop the picture.
   */
  it('writes the row when only an item photograph changed', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    const result = await updateAsset({
      siteUrl: SITE,
      token: 't',
      existing: bulkRow({ units: serialiseUnits([{ index: 0, serialNumber: 'HA2KJDSW' }]) }),
      edits: {
        units: serialiseUnits([
          { index: 0, serialNumber: 'HA2KJDSW', photoUrl: '/sites/it/photos/tab-1.jpg' },
        ]),
      },
    });

    expect(result.changes).toEqual([]);
    expect(rowWrites(sp.calls)).toHaveLength(1);
  });
});

describe('planEdit', () => {
  /** The one place stray spaces come off, so storage never holds two serials
   *  that differ only by one. */
  it('trims the unit values on the way to SharePoint', () => {
    const { record } = planEdit(bulkRow(), {
      units: serialiseUnits([{ index: 0, serialNumber: ' HA2KJDSW ' }]),
    });

    expect(parseUnits(record.units)[0].serialNumber).toBe('HA2KJDSW');
  });
});

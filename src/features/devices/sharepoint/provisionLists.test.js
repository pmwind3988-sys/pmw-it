import {
  describe, it, expect, afterEach, vi,
} from 'vitest';
import { fieldBody, renameBody, provisionLists } from './provisionLists.js';

describe('fieldBody', () => {
  it('creates a text column as a plain SP.Field', () => {
    expect(fieldBody({ StaticName: 'Owner', Title: 'Owner', kind: 'text' })).toEqual({
      __metadata: { type: 'SP.Field' },
      Title: 'Owner', StaticName: 'Owner', FieldTypeKind: 2, Required: false,
    });
  });

  it('creates the column under its internal name, not its display name', () => {
    // SharePoint derives the internal name from the Title a field is created
    // with, so creating this one as 'Device Type' would leave it addressable
    // only as Device_x0020_Type and every item write would fail.
    const body = fieldBody({ StaticName: 'DeviceType', Title: 'Device Type', kind: 'text' });
    expect(body.Title).toBe('DeviceType');
    expect(body.StaticName).toBe('DeviceType');
  });

  it('creates a DateTime column with the time kept', () => {
    const body = fieldBody({ StaticName: 'ScannedOn', Title: 'Scanned On', kind: 'datetime' });
    expect(body.__metadata.type).toBe('SP.FieldDateTime');
    expect(body.FieldTypeKind).toBe(4);
    // 1 = DateTime. 0 would be DateOnly and would throw the time away.
    expect(body.DisplayFormat).toBe(1);
  });

  it('creates a Note column as plain text, not rich text', () => {
    const body = fieldBody({ StaticName: 'RawReport', Title: 'Raw Report', kind: 'note' });
    expect(body.__metadata.type).toBe('SP.FieldMultiLineText');
    expect(body.FieldTypeKind).toBe(3);
    expect(body.RichText).toBe(false);
    expect(body.AppendOnly).toBe(false);
  });

  it('creates a choice column WITH its choices, as SP.FieldChoice', () => {
    const body = fieldBody({
      StaticName: 'DeviceType', Title: 'Device Type', kind: 'choice',
      choices: ['Laptop', 'Desktop', 'Unknown'],
    });
    // Verified against the tenant: sending Choices on the base SP.Field fails
    // with "The property 'Choices' does not exist on type 'SP.Field'".
    expect(body.__metadata.type).toBe('SP.FieldChoice');
    expect(body.FieldTypeKind).toBe(6);
    expect(body.Choices).toEqual({ results: ['Laptop', 'Desktop', 'Unknown'] });
  });

  it('creates a number column with no decimal places', () => {
    const body = fieldBody({ StaticName: 'InstalledRamGB', Title: 'RAM', kind: 'number' });
    expect(body.__metadata.type).toBe('SP.FieldNumber');
    expect(body.FieldTypeKind).toBe(9);
    expect(body.DisplayFormat).toBe(0);
  });

  it('creates a boolean column', () => {
    const body = fieldBody({ StaticName: 'HasHdd', Title: 'Has HDD', kind: 'boolean' });
    expect(body.FieldTypeKind).toBe(8);
  });
});

describe('renameBody', () => {
  it('sets the display name and touches nothing else', () => {
    // Anything beyond Title would ask SharePoint to change what creation
    // already fixed, the internal name most of all.
    expect(renameBody({ StaticName: 'DeviceType', Title: 'Device Type', kind: 'choice' })).toEqual({
      __metadata: { type: 'SP.FieldChoice' },
      Title: 'Device Type',
    });
  });
});

// The provisioning requests themselves. Only `fetch` is faked — spFetch,
// ensureList and ensureColumns all run for real against it.
const SITE = 'https://contoso.sharepoint.com/sites/it';

function fakeSharePoint({ existingFields = [], renameStatus = 200 } = {}) {
  const calls = [];

  const reply = (body, status = 200) => ({
    ok: status >= 200 && status < 300,
    status,
    json: async () => body,
    text: async () => JSON.stringify(body),
  });

  return {
    calls,
    fetch: async (url, init = {}) => {
      const method = init.method ?? 'GET';
      calls.push({
        url,
        method,
        headers: init.headers ?? {},
        body: init.body === undefined ? undefined : JSON.parse(init.body),
      });

      if (url.endsWith('/_api/contextinfo')) {
        return reply({ d: { GetContextWebInformation: { FormDigestValue: 'DIGEST' } } });
      }
      if (url.includes('/fields?$select=')) {
        return reply({
          d: {
            results: existingFields.map(({ internalName, title, staticName }) => ({
              InternalName: internalName,
              StaticName: staticName ?? internalName,
              Title: title,
            })),
          },
        });
      }
      if (url.includes('getByInternalNameOrTitle(')) return reply({}, renameStatus);
      // Both lists already exist, so ensureList returns without creating.
      return reply({});
    },
  };
}

const renamesFor = (calls, staticName) => calls.filter(
  (call) => call.url.includes(`getByInternalNameOrTitle('${staticName}')`),
);

const createFor = (calls, staticName) => calls.find(
  (call) => call.method === 'POST'
    && call.url.endsWith('/fields')
    && call.body?.StaticName === staticName,
);

describe('provisionLists', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('creates a column under its internal name, then renames it for display', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    const created = createFor(sp.calls, 'DeviceType');
    expect(created.body.Title).toBe('DeviceType');

    const renames = renamesFor(sp.calls, 'DeviceType');
    expect(renames).toHaveLength(1);
    // A plain POST to a field endpoint creates a field; only the MERGE
    // override edits the one already there.
    expect(renames[0].headers['X-HTTP-Method']).toBe('MERGE');
    expect(renames[0].body).toEqual({
      __metadata: { type: 'SP.FieldChoice' },
      Title: 'Device Type',
    });
    expect(sp.calls.indexOf(renames[0])).toBeGreaterThan(sp.calls.indexOf(created));
  });

  it('does not spend a request renaming a column to the name it already has', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    expect(createFor(sp.calls, 'Owner')).toBeDefined();
    expect(renamesFor(sp.calls, 'Owner')).toHaveLength(0);
  });

  it('renames a column left showing its internal name as its header', async () => {
    // What a run that died between the create and the rename leaves behind.
    // The next run has to finish the job rather than skip the column.
    const sp = fakeSharePoint({
      existingFields: [{ internalName: 'DeviceType', title: 'DeviceType' }],
    });
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    expect(createFor(sp.calls, 'DeviceType')).toBeUndefined();
    const renames = renamesFor(sp.calls, 'DeviceType');
    expect(renames).toHaveLength(1);
    expect(renames[0].body).toEqual({
      __metadata: { type: 'SP.FieldChoice' },
      Title: 'Device Type',
    });
  });

  it('costs nothing for a column that is already correct', async () => {
    const sp = fakeSharePoint({
      existingFields: [{ internalName: 'DeviceType', title: 'Device Type' }],
    });
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    expect(createFor(sp.calls, 'DeviceType')).toBeUndefined();
    expect(renamesFor(sp.calls, 'DeviceType')).toHaveLength(0);
  });

  it('stops when the display name sits on a stale encoded column', async () => {
    // The pre-fix code's leftover: the header we want, on a field the item
    // writes cannot address. StaticName lies here, which is why the code
    // reads InternalName.
    const sp = fakeSharePoint({
      existingFields: [{
        internalName: 'Device_x0020_Type', staticName: 'DeviceType', title: 'Device Type',
      }],
    });
    vi.stubGlobal('fetch', sp.fetch);

    await expect(provisionLists(SITE, 'token')).rejects.toThrow(/Device_x0020_Type/);
  });

  it('is not fooled by an unrelated column that displays the same header', async () => {
    // A field renamed to our header through the SharePoint UI keeps its own
    // clean internal name, so nothing is stale here. Ours is simply absent.
    const sp = fakeSharePoint({
      existingFields: [{ internalName: 'Notes', title: 'Device Type' }],
    });
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    expect(createFor(sp.calls, 'DeviceType')).toBeDefined();
  });

  it('fails loudly when a rename is rejected', async () => {
    const sp = fakeSharePoint({ renameStatus: 400 });
    vi.stubGlobal('fetch', sp.fetch);

    // Swallowing this would leave the column headed 'OwnerSource' with
    // nothing anywhere to say so.
    await expect(provisionLists(SITE, 'token')).rejects.toThrow(/rename/i);
  });
});

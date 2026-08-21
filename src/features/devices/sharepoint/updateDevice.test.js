import {
  describe, it, expect, afterEach, vi,
} from 'vitest';
import { planEdit, updateDevice, deleteDevice, EDITABLE_FIELDS } from './updateDevice.js';

const SITE = 'https://contoso.sharepoint.com/sites/it';

const row = (overrides) => ({
  id: 7, computerName: 'PC1', owner: 'Ashraf', department: null, deviceType: 'Laptop',
  manualFields: null, ...overrides,
});

describe('EDITABLE_FIELDS', () => {
  it('covers only the values the import guessed', () => {
    expect(EDITABLE_FIELDS).toEqual(['owner', 'department', 'deviceType']);
  });

  it('does not let anything read out of the scan file be retyped', () => {
    for (const field of ['cpuModel', 'installedRamGB', 'windowsVersion', 'storageType']) {
      expect(EDITABLE_FIELDS).not.toContain(field);
    }
  });
});

describe('planEdit', () => {
  it('finds nothing when the value is unchanged', () => {
    expect(planEdit(row(), { owner: 'Ashraf' }).changes).toEqual([]);
  });

  it('records an edit and marks the field as hand-set', () => {
    const result = planEdit(row(), { owner: 'Ashraf Azahari' });
    expect(result.changes).toEqual([
      { fieldName: 'owner', oldValue: 'Ashraf', newValue: 'Ashraf Azahari', changeType: 'Updated' },
    ]);
    expect(result.manualFields).toEqual(['owner']);
  });

  it('calls filling an empty field Added', () => {
    const result = planEdit(row(), { department: 'IT' });
    expect(result.changes[0].changeType).toBe('Added');
    expect(result.manualFields).toEqual(['department']);
  });

  it('hands a field back to the scan file when it is cleared', () => {
    // Blanking is the only way to undo a correction. Without it the field
    // would stay frozen against every future import.
    const result = planEdit(row({ owner: 'Wrong Name', manualFields: ['owner'] }), { owner: '' });
    expect(result.changes[0].changeType).toBe('Removed');
    expect(result.manualFields).toEqual([]);
  });

  it('keeps fields marked earlier that this edit did not touch', () => {
    const result = planEdit(
      row({ manualFields: ['department'] }),
      { owner: 'Ashraf Azahari' },
    );
    expect(result.manualFields.sort()).toEqual(['department', 'owner']);
  });

  it('handles several fields in one edit', () => {
    const result = planEdit(row(), {
      owner: 'Ashraf Azahari', department: 'IT', deviceType: 'Desktop',
    });
    expect(result.changes.map((c) => c.fieldName)).toEqual(['owner', 'department', 'deviceType']);
    expect(result.manualFields).toEqual(['owner', 'department', 'deviceType']);
  });

  it('ignores a field the edit did not include', () => {
    const result = planEdit(row({ department: 'SALES' }), { owner: 'Ashraf Azahari' });
    expect(result.changes.map((c) => c.fieldName)).toEqual(['owner']);
  });
});

function fakeSharePoint({ failOn } = {}) {
  const calls = [];
  const reply = (status = 200) => ({
    ok: status >= 200 && status < 300,
    status,
    json: async () => ({}),
    text: async () => 'boom',
    headers: { get: () => null },
  });

  return {
    calls,
    fetch: async (url, init = {}) => {
      calls.push({ url, method: init.method, headers: init.headers ?? {}, body: init.body ? JSON.parse(init.body) : undefined });
      if (url.endsWith('/_api/contextinfo')) {
        return { ...reply(), json: async () => ({ d: { GetContextWebInformation: { FormDigestValue: 'D' } } }) };
      }
      if (failOn && url.includes(failOn)) return reply(400);
      return reply();
    },
  };
}

const writes = (calls) => calls.filter((c) => c.method === 'POST' && !c.url.endsWith('/contextinfo'));

describe('updateDevice', () => {
  afterEach(() => { vi.unstubAllGlobals(); });

  it('sends only the edited columns plus the manual list', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    await updateDevice({
      siteUrl: SITE, token: 't', existing: row(), edits: { owner: 'Ashraf Azahari' },
      changedBy: 'me@pmw-group.com',
    });

    const patch = writes(sp.calls).find((c) => c.url.includes('items(7)'));
    expect(patch.headers['X-HTTP-Method']).toBe('MERGE');
    expect(patch.body).toEqual({
      ManualFields: 'owner', Owner: 'Ashraf Azahari', OwnerSource: 'Manual',
    });
  });

  it('writes a change row per edited field', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    await updateDevice({
      siteUrl: SITE, token: 't', existing: row(),
      edits: { owner: 'Ashraf Azahari', department: 'IT' },
      changedBy: 'me@pmw-group.com',
    });

    const logged = writes(sp.calls).filter((c) => c.url.includes('Changes'));
    expect(logged.map((c) => c.body.FieldName)).toEqual(['owner', 'department']);
    expect(logged[0].body.ChangedBy).toBe('me@pmw-group.com');
    expect(logged[0].body.ChangedOnMYT).toMatch(/\d{2}\/\d{2}\/\d{4} \d{2}:\d{2} [AP]M/);
  });

  it('does not touch SharePoint at all when nothing changed', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    const result = await updateDevice({
      siteUrl: SITE, token: 't', existing: row(), edits: { owner: 'Ashraf' },
    });

    expect(result.changes).toEqual([]);
    expect(sp.calls).toHaveLength(0);
  });

  it('refuses a row with no id rather than guessing', async () => {
    await expect(updateDevice({
      siteUrl: SITE, token: 't', existing: row({ id: null }), edits: { owner: 'X' },
    })).rejects.toThrow(/no id/);
  });

  it('reports a rejected save instead of pretending it worked', async () => {
    const sp = fakeSharePoint({ failOn: 'items(7)' });
    vi.stubGlobal('fetch', sp.fetch);

    await expect(updateDevice({
      siteUrl: SITE, token: 't', existing: row(), edits: { owner: 'X' },
    })).rejects.toThrow(/Could not save/);
  });
});

describe('deleteDevice', () => {
  afterEach(() => { vi.unstubAllGlobals(); });

  it('deletes the row and records the removal', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    const result = await deleteDevice({
      siteUrl: SITE, token: 't', device: row(), changedBy: 'me@pmw-group.com',
    });

    const removal = writes(sp.calls).find((c) => c.url.includes('items(7)'));
    expect(removal.headers['X-HTTP-Method']).toBe('DELETE');
    expect(removal.headers['IF-MATCH']).toBe('*');

    const logged = writes(sp.calls).find((c) => c.url.includes('Changes'));
    expect(logged.body).toMatchObject({
      Title: 'PC1', ChangeType: 'Removed', OldValue: 'PC1',
    });
    expect(result.removed).toBe('PC1');
  });

  it('refuses a row with no id', async () => {
    await expect(deleteDevice({ siteUrl: SITE, token: 't', device: row({ id: null }) }))
      .rejects.toThrow(/no id/);
  });

  it('does not log a removal when the delete itself failed', async () => {
    const sp = fakeSharePoint({ failOn: 'items(7)' });
    vi.stubGlobal('fetch', sp.fetch);

    await expect(deleteDevice({ siteUrl: SITE, token: 't', device: row() }))
      .rejects.toThrow(/Could not remove/);
    expect(writes(sp.calls).some((c) => c.url.includes('Changes'))).toBe(false);
  });
});

import {
  describe, it, expect, afterEach, vi,
} from 'vitest';
import { provisionLists } from './provisionLists.js';
import { DEVICE_VIEWS } from './deviceViews.js';
import { DEVICE_COLUMNS, CHANGE_COLUMNS } from './deviceSchema.js';

const SITE = 'https://contoso.sharepoint.com/sites/it';

/**
 * Every column already exists, so the run is purely about views. `viewFields`
 * says what each view currently shows, keyed by the path segment that
 * identifies it.
 */
function fakeSharePoint({ existingViews = [], viewFields = {}, viewQueries = {} } = {}) {
  const calls = [];

  const reply = (body, status = 200) => ({
    ok: status >= 200 && status < 300,
    status,
    json: async () => body,
    text: async () => JSON.stringify(body),
  });

  const fields = [...DEVICE_COLUMNS, ...CHANGE_COLUMNS].map((c) => ({
    InternalName: c.StaticName, Title: c.Title,
  }));

  // Anchored on `views/getByTitle` specifically: the list title appears
  // earlier in the same path and would otherwise match first.
  const keyFor = (url) => {
    const path = url.split('?')[0];
    if (path.includes('/defaultView')) {
      return `default:${path.includes('Changes') ? 'changes' : 'devices'}`;
    }
    return decodeURIComponent((/views\/getByTitle\('([^']+)'\)/.exec(path) ?? [])[1] ?? '');
  };

  return {
    calls,
    fetch: async (url, init = {}) => {
      const method = init.method ?? 'GET';
      calls.push({ url, method, body: init.body ? JSON.parse(init.body) : undefined });

      if (url.endsWith('/_api/contextinfo')) {
        return reply({ d: { GetContextWebInformation: { FormDigestValue: 'DIGEST' } } });
      }
      if (url.includes('/fields?$select=')) return reply({ d: { results: fields } });
      if (url.includes('/views?$select=Title')) {
        return reply({ d: { results: existingViews.map((Title) => ({ Title })) } });
      }
      if (url.includes('?$select=ViewQuery')) {
        return reply({ d: { ViewQuery: viewQueries[keyFor(url)] ?? '' } });
      }
      if (url.endsWith('/viewfields') && method === 'GET') {
        const current = viewFields[keyFor(url)];
        return current
          ? reply({ d: { Items: { results: current } } })
          : reply({ d: { Items: { results: [] } } });
      }
      return reply({});
    },
  };
}

const viewCalls = (calls) => calls.filter((c) => c.url.includes('viewfields'));
const addedTo = (calls, marker) => calls
  .filter((c) => c.url.includes(marker) && c.url.includes('addviewfield'))
  .map((c) => decodeURIComponent(/addviewfield\('([^']+)'\)/.exec(c.url)[1]));

describe('provisionLists — views', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('creates the named views and leaves the default one alone', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    const created = sp.calls
      .filter((c) => c.method === 'POST' && c.url.endsWith('/views'))
      .map((c) => c.body.Title);

    expect(created).toEqual(['Needs attention', 'Upgrade candidates']);
    // The built-in view is addressed through /defaultView, never created.
    expect(sp.calls.some((c) => c.url.includes('/defaultView/viewfields'))).toBe(true);
  });

  it('sets the default view fields in the declared order', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    const expected = DEVICE_VIEWS.find((v) => v.isDefault && v.list.includes('Device List')).fields;
    const added = addedTo(sp.calls, "Device%20List')/defaultView");
    expect(added).toEqual(expected);
  });

  it('clears a view before filling it, or the defaults would linger', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    const forDefault = viewCalls(sp.calls)
      .filter((c) => c.url.includes("Device%20List')/defaultView"));
    expect(forDefault[1].url).toContain('removeallviewfields');
    expect(forDefault[2].url).toContain('addviewfield');
  });

  it('does not rewrite a view whose fields are already right', async () => {
    const deviceDefault = DEVICE_VIEWS.find(
      (v) => v.isDefault && v.list.includes('Device List'),
    );
    const changeDefault = DEVICE_VIEWS.find((v) => v.isDefault && v.list.includes('Changes'));
    const named = DEVICE_VIEWS.filter((v) => !v.isDefault);

    const sp = fakeSharePoint({
      existingViews: named.map((v) => v.title),
      viewFields: {
        'default:devices': deviceDefault.fields,
        'default:changes': changeDefault.fields,
        ...Object.fromEntries(named.map((v) => [v.title, v.fields])),
      },
    });
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    // Reads to compare, but no writes at all.
    expect(sp.calls.some((c) => c.url.includes('removeallviewfields'))).toBe(false);
    expect(sp.calls.some((c) => c.url.includes('addviewfield'))).toBe(false);
    expect(sp.calls.some((c) => c.method === 'POST' && c.url.endsWith('/views'))).toBe(false);
  });

  it('repairs a view somebody has since edited', async () => {
    const sp = fakeSharePoint({
      existingViews: ['Needs attention', 'Upgrade candidates'],
      viewFields: { 'default:devices': ['LinkTitle'] },
    });
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    expect(addedTo(sp.calls, "Device%20List')/defaultView").length).toBeGreaterThan(1);
  });

  it('builds the views only after the columns they show exist', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    const lastField = sp.calls.map((c) => c.url).reduce(
      (acc, url, i) => (url.endsWith('/fields') ? i : acc), -1,
    );
    const firstView = sp.calls.findIndex((c) => c.url.includes('viewfields'));
    expect(firstView).toBeGreaterThan(lastField);
  });
});

describe('provisionLists — view filters and sorts', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  const merges = (calls) => calls.filter(
    (c) => c.method === 'POST' && c.body?.ViewQuery !== undefined,
  );

  it('merges the sort onto a default view, which cannot take one at creation', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    const changeDefault = DEVICE_VIEWS.find((v) => v.isDefault && v.list.includes('Changes'));
    const onDefault = merges(sp.calls).filter((c) => c.url.includes('/defaultView'));

    expect(onDefault.some((c) => c.body.ViewQuery === changeDefault.query)).toBe(true);
  });

  it('leaves a view query alone when it already matches', async () => {
    const named = DEVICE_VIEWS.filter((v) => !v.isDefault);
    const changeDefault = DEVICE_VIEWS.find((v) => v.isDefault && v.list.includes('Changes'));

    const sp = fakeSharePoint({
      existingViews: named.map((v) => v.title),
      viewQueries: {
        'default:changes': changeDefault.query,
        ...Object.fromEntries(named.map((v) => [v.title, v.query])),
      },
    });
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    expect(merges(sp.calls)).toHaveLength(0);
  });

  it('does not merge a query onto a view that declares none', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    await provisionLists(SITE, 'token');

    // The devices default view is deliberately unfiltered: it is the whole
    // register, and a filter there would hide machines without saying so.
    const onDevicesDefault = merges(sp.calls)
      .filter((c) => c.url.includes("Device%20List')/defaultView"));
    expect(onDevicesDefault).toHaveLength(0);
  });
});

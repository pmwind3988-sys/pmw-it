import { describe, it, expect, vi, afterEach } from 'vitest';
import { searchPeople, normalisePerson, MIN_QUERY } from './peopleSearch.js';

const SITE = 'https://example.sharepoint.com/sites/IT';

const picker = (entries) => ({
  ok: true,
  status: 200,
  json: async () => ({ d: { ClientPeoplePickerSearchUser: JSON.stringify(entries) } }),
});

afterEach(() => { vi.unstubAllGlobals(); });

describe('normalisePerson', () => {
  it('reads a directory user', () => {
    expect(normalisePerson({
      Key: 'i:0#.f|membership|amir@pmw.com',
      DisplayText: 'Amir Hakim',
      EntityData: { Email: 'Amir@PMW.com', Title: 'Engineer', Department: 'Engineering' },
    })).toEqual({
      name: 'Amir Hakim',
      email: 'amir@pmw.com',
      login: 'i:0#.f|membership|amir@pmw.com',
      title: 'Engineer',
      department: 'Engineering',
    });
  });

  /** The picker puts the address in Description for an email-matched entry. */
  it('falls back to the description when there is no EntityData', () => {
    expect(normalisePerson({
      Key: 'i:0#.f|membership|evonne@pmw.com',
      DisplayText: 'Evonne',
      Description: 'evonne@pmw.com',
    }).email).toBe('evonne@pmw.com');
  });

  /** And to the claim itself, which is the only thing a bare member claim has. */
  it('falls back to the login claim', () => {
    expect(normalisePerson({
      Key: 'i:0#.f|membership|carmen@pmw.com',
      DisplayText: 'Carmen',
    }).email).toBe('carmen@pmw.com');
  });

  it('never comes back with no name at all', () => {
    expect(normalisePerson({ Key: 'i:0#.f|membership|x@pmw.com' }).name).toBe('x@pmw.com');
  });

  it('is safe on nothing', () => {
    expect(normalisePerson(undefined).email).toBe('');
  });
});

describe('searchPeople', () => {
  it('finds people', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => picker([
      { Key: 'i:0#.f|membership|amir@pmw.com', DisplayText: 'Amir', EntityData: { Email: 'amir@pmw.com' } },
    ])));

    const people = await searchPeople(SITE, 'token', 'ami');
    expect(people).toHaveLength(1);
    expect(people[0].name).toBe('Amir');
  });

  /** Two letters find half the company. */
  it('does not search on fewer than three characters, or ask the network to', async () => {
    const fetcher = vi.fn();
    vi.stubGlobal('fetch', fetcher);

    expect(await searchPeople(SITE, 'token', 'am')).toEqual([]);
    expect(await searchPeople(SITE, 'token', '   ')).toEqual([]);
    expect(fetcher).not.toHaveBeenCalled();
  });

  /** Without an email there is nothing to key a person's items on. */
  it('drops an entry it cannot identify', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => picker([
      { Key: 'c:0(.s|true', DisplayText: 'Everyone' },
      { Key: 'i:0#.f|membership|amir@pmw.com', DisplayText: 'Amir' },
    ])));

    const people = await searchPeople(SITE, 'token', 'eve');
    expect(people).toHaveLength(1);
    expect(people[0].email).toBe('amir@pmw.com');
  });

  /** The picker answers with JSON inside its JSON. */
  it('copes when the inner payload is not a string', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => ({
      ok: true,
      status: 200,
      json: async () => ({
        d: {
          ClientPeoplePickerSearchUser: [
            { Key: 'i:0#.f|membership|amir@pmw.com', DisplayText: 'Amir' },
          ],
        },
      }),
    })));

    expect(await searchPeople(SITE, 'token', 'ami')).toHaveLength(1);
  });

  it('returns nothing rather than throwing on unreadable results', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => ({
      ok: true,
      status: 200,
      json: async () => ({ d: { ClientPeoplePickerSearchUser: 'not json at all' } }),
    })));

    expect(await searchPeople(SITE, 'token', 'ami')).toEqual([]);
  });

  it('throws when the directory itself refuses', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => ({ ok: false, status: 403, json: async () => ({}) })));
    await expect(searchPeople(SITE, 'token', 'ami')).rejects.toThrow(/403/);
  });

  it('asks only for users, never for groups', async () => {
    const fetcher = vi.fn(async () => picker([]));
    vi.stubGlobal('fetch', fetcher);

    await searchPeople(SITE, 'token', 'ami');
    const body = JSON.parse(fetcher.mock.calls[0][1].body);

    expect(body.queryParams.PrincipalType).toBe(1);
    expect(body.queryParams.QueryString).toBe('ami');
  });

  it('agrees with the minimum it documents', () => {
    expect(MIN_QUERY).toBe(3);
  });
});

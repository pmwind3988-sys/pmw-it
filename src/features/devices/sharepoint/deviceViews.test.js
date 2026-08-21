import { describe, it, expect } from 'vitest';
import { DEVICE_VIEWS } from './deviceViews.js';
import {
  DEVICE_COLUMNS, CHANGE_COLUMNS, DEVICE_LIST_NAME, CHANGE_LIST_NAME,
} from './deviceSchema.js';

/** SharePoint's own fields, which exist on every list without being created. */
const BUILT_IN = new Set(['LinkTitle', 'Title', 'ID', 'Created', 'Modified', 'Author', 'Editor']);

const columnsFor = (list) => new Set([
  ...(list === DEVICE_LIST_NAME ? DEVICE_COLUMNS : CHANGE_COLUMNS).map((c) => c.StaticName),
  ...BUILT_IN,
]);

/** Every `Name="..."` referenced inside a view's CAML query. */
const fieldsInQuery = (query) =>
  [...(query ?? '').matchAll(/Name="([^"]+)"/g)].map((m) => m[1]);

describe('DEVICE_VIEWS', () => {
  it('covers both lists', () => {
    const lists = new Set(DEVICE_VIEWS.map((v) => v.list));
    expect(lists).toEqual(new Set([DEVICE_LIST_NAME, CHANGE_LIST_NAME]));
  });

  it('gives each list exactly one default view', () => {
    for (const list of [DEVICE_LIST_NAME, CHANGE_LIST_NAME]) {
      const defaults = DEVICE_VIEWS.filter((v) => v.list === list && v.isDefault);
      expect(defaults).toHaveLength(1);
    }
  });

  it('has no duplicate view titles within a list', () => {
    for (const list of [DEVICE_LIST_NAME, CHANGE_LIST_NAME]) {
      const titles = DEVICE_VIEWS.filter((v) => v.list === list).map((v) => v.title);
      expect(new Set(titles).size).toBe(titles.length);
    }
  });

  // The one that matters: a typo here creates a view silently missing a
  // column, and nothing downstream would ever complain.
  it('only ever names columns that exist on the list it belongs to', () => {
    for (const view of DEVICE_VIEWS) {
      const known = columnsFor(view.list);
      for (const field of view.fields) {
        expect(known, `${view.title} → ${field}`).toContain(field);
      }
    }
  });

  it('only filters and sorts on columns that exist', () => {
    for (const view of DEVICE_VIEWS) {
      const known = columnsFor(view.list);
      for (const field of fieldsInQuery(view.query)) {
        expect(known, `${view.title} query → ${field}`).toContain(field);
      }
    }
  });

  it('leads every view with the clickable name', () => {
    for (const view of DEVICE_VIEWS) {
      expect(view.fields[0]).toBe('LinkTitle');
    }
  });

  it('keeps the multi-line columns out of the general views', () => {
    // A single Raw Report cell is taller than the screen.
    const heavy = ['RawReport', 'RamSlotInfoRaw', 'StorageDrivesRaw', 'EmailDataFiles',
      'ServerFolders', 'ServerCredentials', 'AntivirusProducts', 'MonitorsRaw', 'ExtraFields'];

    for (const view of DEVICE_VIEWS) {
      for (const field of heavy) {
        expect(view.fields, `${view.title}`).not.toContain(field);
      }
    }
  });

  it('shows Risk Reasons only where the list is already narrowed to it', () => {
    const withReasons = DEVICE_VIEWS.filter((v) => v.fields.includes('RiskReasons'));
    expect(withReasons.map((v) => v.title)).toEqual(['Needs attention']);
    expect(withReasons[0].query).toContain('RiskLevel');
  });

  it('sorts the change log newest first', () => {
    const changes = DEVICE_VIEWS.find((v) => v.list === CHANGE_LIST_NAME);
    expect(changes.query).toContain('Ascending="FALSE"');
    expect(changes.query).toContain('ChangedOn');
  });
});

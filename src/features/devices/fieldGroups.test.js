import { describe, it, expect } from 'vitest';
import { FIELD_GROUPS, ALL_KEYS, groupsFor, labelFor } from './fieldGroups.js';

describe('fieldGroups', () => {
  it('places every schema field in exactly one group', () => {
    const placed = FIELD_GROUPS.flatMap((group) => group.keys);
    expect([...placed].sort()).toEqual([...ALL_KEYS].sort());
  });

  it('never invents an "Other" group while the schema is fully covered', () => {
    expect(FIELD_GROUPS.some((group) => group.id === 'other')).toBe(false);
  });

  it('leaves the raw report out of the groups', () => {
    expect(FIELD_GROUPS.flatMap((g) => g.keys)).not.toContain('rawReport');
  });

  it('labels a field the way the list column is labelled', () => {
    expect(labelFor('installedRamGB')).toBe('Installed RAM (GB)');
    expect(labelFor('computerName')).toBe('Computer');
  });

  it('drops the groups a device has nothing in', () => {
    const groups = groupsFor({ computerName: 'ASHRAF-PC', installedRamGB: 8 });
    expect(groups.map((g) => g.id)).toEqual(['identity', 'memory']);
    expect(groups[0].fields.map((f) => f.key)).toEqual(['computerName']);
  });

  it('keeps the empty ones when asked', () => {
    const groups = groupsFor({ computerName: 'A' }, { includeEmpty: true });
    expect(groups).toHaveLength(FIELD_GROUPS.length);
  });
});

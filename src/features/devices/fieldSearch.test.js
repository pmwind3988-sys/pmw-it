import { describe, it, expect } from 'vitest';
import { matchingGroups } from './fieldSearch.js';

const groups = [
  { id: 'network', title: 'Network', fields: [{ key: 'ipAddress', label: 'IP address' }, { key: 'ssid', label: 'Wi-Fi name' }] },
  { id: 'software', title: 'Software', fields: [{ key: 'adobeProducts', label: 'Adobe products' }] },
];
const values = { ipAddress: '10.1.4.87', ssid: 'PMW-Staff', adobeProducts: 'Acrobat Reader' };
const valueOf = (field) => values[field.key];

describe('matchingGroups', () => {
  it('returns every group untouched for an empty search', () => {
    expect(matchingGroups(groups, '   ', valueOf)).toBe(groups);
  });

  it('matches a label, ignoring case', () => {
    const out = matchingGroups(groups, 'WI-FI', valueOf);
    expect(out).toEqual([{ ...groups[0], fields: [groups[0].fields[1]] }]);
  });

  it('matches a value', () => {
    const out = matchingGroups(groups, '10.1.4', valueOf);
    expect(out.map((g) => g.id)).toEqual(['network']);
    expect(out[0].fields.map((f) => f.key)).toEqual(['ipAddress']);
  });

  it('drops groups with nothing matching', () => {
    expect(matchingGroups(groups, 'nothing like this', valueOf)).toEqual([]);
  });
});

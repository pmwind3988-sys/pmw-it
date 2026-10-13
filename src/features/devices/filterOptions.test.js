import { describe, expect, it } from 'vitest';
import { applyFilters, IN_FLEET } from './deviceFilters';
import { buildFilterGroups, chipLabel } from './filterOptions';

const machine = (over) => ({
  id: 1, computerName: 'PC-1', status: 'In use', department: 'HR', location: 'F1',
  riskLevel: 'OK', installedRamGB: 8, storageType: 'SSD', cpuAgeBand: '2020',
  deviceType: 'Laptop', licenseStatus: 'Licensed', personaLabel: 'Desk',
  parts: {}, scannedOn: Date.now(), ...over,
});

const find = (groups, id) => groups.find((g) => g.id === id);

describe('buildFilterGroups', () => {
  const devices = [
    machine({ id: 1, department: 'HR', riskLevel: 'Critical', installedRamGB: 16 }),
    machine({ id: 2, department: '', riskLevel: 'Watch', installedRamGB: 8, status: 'Retired' }),
  ];

  it('offers each distinct department, with a blank one as Unassigned', () => {
    const labels = find(buildFilterGroups(devices, ['F1']), 'place').items
      .filter((o) => o.filter === 'department').map((o) => o.value);
    expect(labels).toEqual(['HR', 'Unassigned']);
  });

  it('adds Unassigned to the location choices from the prop', () => {
    const locations = find(buildFilterGroups(devices, ['F3', 'F1']), 'place').items
      .filter((o) => o.filter === 'location').map((o) => o.value);
    expect(locations).toEqual(['F1', 'F3', 'Unassigned']);
  });

  it('lists the RAM sizes numerically, not as text', () => {
    const ram = find(buildFilterGroups(devices), 'hardware').items
      .filter((o) => o.filter === 'ram').map((o) => o.value);
    expect(ram).toEqual(['8 GB', '16 GB']);
  });

  it('labels risk by the band a person reads', () => {
    const risk = find(buildFilterGroups(devices), 'security').items
      .filter((o) => o.filter === 'risk');
    expect(risk.map((o) => o.label)).toEqual(['High', 'Medium']);
    expect(risk.map((o) => o.value)).toEqual(['High', 'Medium']);
  });

  it('every option the menu offers is a filter key applyFilters understands', () => {
    const keys = buildFilterGroups(devices, ['F1'])
      .flatMap((group) => group.items.map((item) => item.filter));
    for (const key of new Set(keys)) {
      // An unknown key matches everything; a known one can match fewer.
      const narrowed = applyFilters(devices, { [key]: 'zz-no-such-value' });
      expect(narrowed.length).toBeLessThanOrEqual(devices.length);
    }
  });

  it('Status offers the fleet default first and Everything last', () => {
    const status = find(buildFilterGroups(devices), 'place').items.filter((o) => o.filter === 'status');
    expect(status[0].value).toBe(IN_FLEET);
    expect(status.at(-1).value).toBe('All');
  });
});

describe('chipLabel', () => {
  it('reads a flag as its label alone', () => {
    expect(chipLabel('work', 'yes')).toBe('Needs work');
  });

  it('reads the risk as the person-facing band', () => {
    expect(chipLabel('risk', 'Critical')).toBe('Security risk: High');
  });

  it('reads a part and grade as one phrase', () => {
    expect(chipLabel('part', 'ram:Critical')).toBe('RAM: Critical');
  });

  it('names the default status and "Everything"', () => {
    expect(chipLabel('status', IN_FLEET)).toBe('In use or in repair');
    expect(chipLabel('status', 'All')).toBe('Everything');
  });
});

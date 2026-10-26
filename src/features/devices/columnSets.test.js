import { describe, expect, it } from 'vitest';
import {
  ALL_COLUMNS, COLUMN_SETS, COLUMN_SET_IDS, DEFAULT_SET, columnsForKeys, columnsForSet,
  keysForSet, parseColumnSet, setMatching, toggleColumn,
} from './columnSets';

const allKeys = new Set(ALL_COLUMNS.map((column) => column.key));

describe('column sets', () => {
  it('every key a set names is a real column', () => {
    for (const id of COLUMN_SET_IDS) {
      const missing = COLUMN_SETS[id].keys.filter((key) => !allKeys.has(key));
      expect(missing, `set "${id}"`).toEqual([]);
    }
  });

  it('Everything draws every column except the frozen name', () => {
    const everything = columnsForSet('everything').map((column) => column.key);
    expect(everything).toEqual(ALL_COLUMNS.map((c) => c.key).filter((k) => k !== 'computerName'));
  });

  it('no set repeats a column', () => {
    for (const id of COLUMN_SET_IDS) {
      const keys = COLUMN_SETS[id].keys;
      expect(new Set(keys).size, `set "${id}"`).toBe(keys.length);
    }
  });

  it('Essentials is the default and shows the five grades and Scanned', () => {
    expect(DEFAULT_SET).toBe('essentials');
    const keys = COLUMN_SETS.essentials.keys;
    for (const grade of ['gradeCpu', 'gradeRam', 'gradeStorage', 'gradeGraphics', 'gradeWindows', 'scannedOn']) {
      expect(keys).toContain(grade);
    }
  });

  it('Security names the risk column as Security risk', () => {
    const risk = columnsForSet('security').find((column) => column.key === 'riskLevel');
    expect(risk.label).toBe('Security risk');
  });

  it('hiding a column takes it out and keeps the rest in order', () => {
    const keys = ['status', 'owner', 'scannedOn'];
    expect(toggleColumn(keys, 'owner')).toEqual(['status', 'scannedOn']);
  });

  it('showing a column puts it where the full order has it, not after Scanned', () => {
    const keys = COLUMN_SETS.essentials.keys;
    const next = toggleColumn(keys, 'licenseStatus');
    expect(next.indexOf('licenseStatus')).toBeLessThan(next.indexOf('scannedOn'));
    expect(next.length).toBe(keys.length + 1);
  });

  it('a column earlier than everything shown goes first', () => {
    expect(toggleColumn(['scannedOn'], 'status')).toEqual(['status', 'scannedOn']);
  });

  it('the computer name and unknown keys are not choices', () => {
    const keys = ['status'];
    expect(toggleColumn(keys, 'computerName')).toBe(keys);
    expect(toggleColumn(keys, 'nonsense')).toBe(keys);
  });

  it('names the set a pick matches, and none once it is changed', () => {
    expect(setMatching(keysForSet('hardware'))).toBe('hardware');
    expect(setMatching(toggleColumn(keysForSet('hardware'), 'owner'))).toBeNull();
    expect(columnsForKeys(['owner', 'nonsense']).map((c) => c.key)).toEqual(['owner']);
  });

  it('an unknown or missing stored choice falls back to Essentials', () => {
    expect(parseColumnSet('hardware')).toBe('hardware');
    expect(parseColumnSet('nonsense')).toBe(DEFAULT_SET);
    expect(parseColumnSet(null)).toBe(DEFAULT_SET);
  });
});

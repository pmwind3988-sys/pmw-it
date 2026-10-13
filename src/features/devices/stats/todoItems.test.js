import { describe, it, expect } from 'vitest';
import { todoItems } from './todoItems.js';

const NOW = Date.UTC(2026, 9, 13);
const DAY = 86_400_000;

const base = {
  riskLevel: 'OK', osSupported: true, avProtected: true, licenseStatus: 'Authentic',
  serverDependent: false, formFactorMatches: true, criticalParts: [], scannedOn: NOW - DAY,
  scanComplete: true,
};

const devices = [
  { ...base, id: 1, criticalParts: ['cpu'] },
  { ...base, id: 2, osSupported: false },
  { ...base, id: 3, osSupported: false, avProtected: false },
  { ...base, id: 4, licenseStatus: 'Unlicensed', formFactorMatches: false },
  { ...base, id: 5, serverDependent: true, networkRisk: 'Severe' },
  { ...base, id: 6, scannedOn: NOW - 200 * DAY },
];

describe('todoItems', () => {
  const items = todoItems(devices, NOW);

  it('counts each problem the way the List filter counts it', () => {
    const counts = Object.fromEntries(items.map((item) => [item.id, item.count]));
    expect(counts).toEqual({
      'critical-part': 1,
      'unsupported-windows': 2,
      'no-antivirus': 1,
      'unlicensed-office': 1,
      'server-over-wifi': 1,
      'form-factor': 1,
      'stale-scan': 1,
    });
  });

  it('orders by severity first, then by count, worst first', () => {
    expect(items.map((item) => item.severity)).toEqual([
      'urgent', 'urgent', 'urgent', 'warning', 'warning', 'suggestion', 'upkeep',
    ]);
    expect(items[0].id).toBe('unsupported-windows');
  });

  it('puts a larger count first within the same severity', () => {
    const many = [
      ...devices,
      { ...base, id: 7, criticalParts: ['ram'] },
      { ...base, id: 8, criticalParts: ['ram'] },
    ];
    const urgent = todoItems(many, NOW).filter((item) => item.severity === 'urgent');
    expect(urgent.map((item) => item.id)).toEqual(['critical-part', 'unsupported-windows', 'no-antivirus']);
    expect(urgent[0].count).toBe(3);
  });

  it('leaves out every item with nothing to do', () => {
    expect(todoItems([{ ...base, id: 9 }], NOW)).toEqual([]);
  });

  it('carries the filter each card opens', () => {
    expect(items.find((item) => item.id === 'no-antivirus').filter).toEqual({ key: 'av', value: 'Unprotected' });
    expect(items.find((item) => item.id === 'stale-scan').filter).toEqual({ key: 'stale', value: '1' });
  });

  it('does not offer a card for a part that only needs attention', () => {
    const attention = [{ ...base, id: 10, parts: { cpu: { grade: 'Needs attention' } } }];
    expect(todoItems(attention, NOW)).toEqual([]);
  });
});

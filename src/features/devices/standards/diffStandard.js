import { PROFILE_KEYS, PROFILE_SHORT } from './defaultStandard.js';

const CUT_PARTS = { cpu: 'CPU', ram: 'RAM', storageSize: 'storage size' };
const CUTS = { criticalBelow: 'critical below', attentionBelow: 'needs attention below', optimalFrom: 'optimal from' };
const CHOICE_PARTS = ['storageType', 'graphics', 'windows'];
const CHOICE_NAME = {
  'HDD only': 'HDD only', Mixed: 'mixed disks', 'SSD only': 'SSD only',
  dedicated: 'dedicated card', builtIn: 'built-in graphics',
  win11: 'Windows 11', win10Supported: 'Windows 10 (supported)', outOfSupport: 'Windows out of support', other: 'other Windows',
};

/** Every difference between two standards, one readable line each, in page order. */
export function diffStandard(before, after) {
  const lines = [];
  for (const key of PROFILE_KEYS) {
    const a = before.profiles[key];
    const b = after.profiles[key];
    const who = PROFILE_SHORT[key];
    for (const [part, name] of Object.entries(CUT_PARTS)) {
      for (const [cut, cutName] of Object.entries(CUTS)) {
        if (a[part][cut] !== b[part][cut]) lines.push(`${who} ${name} ${cutName} ${a[part][cut]} → ${b[part][cut]}`);
      }
    }
    for (const part of CHOICE_PARTS) {
      for (const kind of Object.keys(b[part])) {
        if (a[part][kind] !== b[part][kind]) lines.push(`${who} ${CHOICE_NAME[kind]} ${a[part][kind]} → ${b[part][kind]}`);
      }
    }
  }
  const departments = new Set([...Object.keys(before.departments), ...Object.keys(after.departments)]);
  for (const department of [...departments].sort()) {
    if (before.departments[department] !== after.departments[department]) {
      lines.push(`${department} judged as ${PROFILE_SHORT[after.departments[department]] ?? 'Desk'}`);
    }
  }
  for (const grade of Object.keys(after.colors)) {
    if (String(before.colors[grade]).toLowerCase() !== String(after.colors[grade]).toLowerCase()) {
      lines.push(`${grade} colour changed`);
    }
  }
  return lines;
}

export const summaryOf = (lines) => (lines.length ? lines.join('; ') : 'No changes');

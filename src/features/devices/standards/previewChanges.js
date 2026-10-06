import { PARTS } from './defaultStandard.js';
import { partGrades } from '../derive/partGrades.js';

/** The machines a new standard would move, per part and grade change, biggest moves first per part. */
export function previewChanges(devices, before, after) {
  const counts = new Map();
  for (const device of devices) {
    const was = partGrades(device, before).parts;
    const will = partGrades(device, after).parts;
    for (const { key } of PARTS) {
      if (was[key].grade === will[key].grade) continue;
      const id = `${key}|${was[key].grade}|${will[key].grade}`;
      counts.set(id, (counts.get(id) ?? 0) + 1);
    }
  }
  const order = PARTS.map((part) => part.key);
  return [...counts]
    .map(([id, count]) => { const [part, from, to] = id.split('|'); return { part, from, to, count }; })
    .sort((a, b) => order.indexOf(a.part) - order.indexOf(b.part) || b.count - a.count);
}

import { parseRamSlots, parseSize } from '../parse/parseValues.js';

const mostCommon = (values) => {
  const counts = new Map();
  for (const value of values) counts.set(value, (counts.get(value) ?? 0) + 1);

  let winner = null;
  let best = 0;
  for (const [value, count] of counts) {
    if (count > best) {
      winner = value;
      best = count;
    }
  }
  return winner;
};

/**
 * `Total RAM` is what Windows reports as USABLE, so an integrated GPU's
 * reserved share is missing from it: a 16 GB laptop reports 15 GB and an 8 GB
 * one reports 7 GB. Ranking machines on that figure puts a 16 GB laptop below
 * an 8 GB one, so the sum of the sticks is the authoritative number and the
 * reported figure is kept only to explain the difference.
 */
export function deriveRam(slotLines, totalRamLines) {
  const { sticks, totalSlots, usedSlots } = parseRamSlots(slotLines);

  const sizes = sticks.map((stick) => stick.sizeGB).filter((n) => typeof n === 'number');
  const installedRamGB = sizes.length ? sizes.reduce((a, b) => a + b, 0) : null;
  const reportedRamGB = totalRamLines.length ? parseSize(totalRamLines[0]) : null;

  const speeds = sticks.map((stick) => stick.speedMhz).filter((n) => typeof n === 'number');
  const types = sticks.map((stick) => stick.type).filter(Boolean);

  return {
    installedRamGB,
    reportedRamGB,
    ramDiscrepancy:
      installedRamGB !== null && reportedRamGB !== null && installedRamGB !== reportedRamGB,
    ramType: types.length ? mostCommon(types) : 'Unknown',
    // Mixed sticks run at the slowest module's speed.
    ramSpeedMhz: speeds.length ? Math.min(...speeds) : null,
    ramSlotsUsed: usedSlots,
    ramSlotsTotal: totalSlots,
    ramUpgradable:
      typeof totalSlots === 'number' && typeof usedSlots === 'number' && usedSlots < totalSlots,
  };
}

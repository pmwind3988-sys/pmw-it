import { inFleet, statusOf, SPARE, RETIRED } from '../lifecycle/status.js';
import { labelOf, UNASSIGNED } from '../deviceFilters.js';
import { cleanLocation } from './locations.js';

/**
 * Every number a tile on the map shows, at every level -- location,
 * department, IT Stash, Graveyard -- computed from the same rows the register
 * reads, so a tile and the list it opens cannot disagree.
 */
export const PLACES = { STASH: 'stash', GRAVEYARD: 'graveyard', NO_LOCATION: 'nolocation' };

const LEVELS = ['Critical', 'Needs Attention', 'Moderate', 'Optimal', 'Unknown'];

export function summarise(name, devices) {
  const fit = Object.fromEntries(LEVELS.map((level) => [level, 0]));
  let laptops = 0;
  let desktops = 0;
  let other = 0;

  for (const device of devices) {
    fit[LEVELS.includes(device.fitStatus) ? device.fitStatus : 'Unknown'] += 1;
    if (device.deviceType === 'Laptop') laptops += 1;
    else if (device.deviceType === 'Desktop') desktops += 1;
    else other += 1;
  }

  const count = devices.length;
  return {
    name,
    count,
    laptops,
    desktops,
    other,
    fit,
    critical: fit.Critical,
    attention: fit['Needs Attention'],
    health: count
      ? LEVELS.map((level) => ({ level, share: fit[level] / count })).filter((part) => part.share > 0)
      : [],
  };
}

function groupBy(items, keyOf) {
  const groups = new Map();
  for (const item of items) {
    const key = keyOf(item);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(item);
  }
  return groups;
}

const largestFirst = (a, b) => b.count - a.count || a.name.localeCompare(b.name);
const departmentOf = (device) => labelOf(device.department);

export function worldTiles(devices) {
  const fleet = devices.filter(inFleet);
  const located = groupBy(fleet.filter((d) => cleanLocation(d.location)), (d) => cleanLocation(d.location));
  const unlocated = fleet.filter((d) => !cleanLocation(d.location));

  const locations = [...located].map(([code, rows]) => {
    const chips = [...groupBy(rows, departmentOf)]
      .map(([name, list]) => ({ name, count: list.length }))
      .sort(largestFirst);
    return { ...summarise(code, rows), code, departments: chips.length, chips };
  }).sort(largestFirst);

  return {
    locations,
    noLocation: unlocated.length ? summarise('No location yet', unlocated) : null,
    stash: summarise('IT Stash', devices.filter((d) => statusOf(d) === SPARE)),
    graveyard: summarise('Graveyard', devices.filter((d) => statusOf(d) === RETIRED)),
  };
}

export function locationTiles(devices, code) {
  const wanted = cleanLocation(code);
  const rows = devices.filter(inFleet).filter((d) => cleanLocation(d.location) === wanted);
  const all = [...groupBy(rows, departmentOf)].map(([name, list]) => summarise(name, list)).sort(largestFirst);
  return {
    centre: summarise(wanted, rows),
    departments: all.filter((tile) => tile.name !== UNASSIGNED),
    unassigned: all.find((tile) => tile.name === UNASSIGNED) ?? null,
  };
}

const FIT_RANK = { Critical: 0, 'Needs Attention': 1, Moderate: 2, Optimal: 3 };

export function machinesIn(devices, { location, department, place } = {}) {
  let rows;
  if (place === PLACES.STASH) rows = devices.filter((d) => statusOf(d) === SPARE);
  else if (place === PLACES.GRAVEYARD) rows = devices.filter((d) => statusOf(d) === RETIRED);
  else if (place === PLACES.NO_LOCATION) rows = devices.filter(inFleet).filter((d) => !cleanLocation(d.location));
  else {
    const wanted = cleanLocation(location);
    rows = devices.filter(inFleet)
      .filter((d) => cleanLocation(d.location) === wanted && departmentOf(d) === department);
  }

  return [...rows].sort((a, b) =>
    (FIT_RANK[a.fitStatus] ?? 4) - (FIT_RANK[b.fitStatus] ?? 4)
    || String(a.computerName ?? '').localeCompare(String(b.computerName ?? '')));
}

export function searchMachines(rows, query) {
  const needle = String(query ?? '').trim().toLowerCase();
  if (!needle) return rows;
  return rows.filter((d) => `${d.owner ?? ''} ${d.computerName ?? ''} ${d.serialNumber ?? ''}`.toLowerCase().includes(needle));
}

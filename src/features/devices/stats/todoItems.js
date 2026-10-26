import { applyFilters, isStale } from '../deviceFilters.js';
import { inFleet } from '../lifecycle/status.js';

/**
 * The To do list: every problem the fleet has, worst first. Each count is the
 * number of machines the List shows when the item's filter is applied, so the
 * number a card prints and the rows it opens cannot disagree.
 */

export const SEVERITY_ORDER = ['urgent', 'warning', 'suggestion', 'upkeep'];

const ITEMS = [
  {
    id: 'critical-part',
    title: 'A critical part',
    why: 'Too slow for the work its desk does',
    severity: 'urgent',
    filter: { key: 'critical', value: '1' },
    count: (devices) => applyFilters(devices, { critical: '1' }).length,
  },
  {
    id: 'unsupported-windows',
    title: 'Windows no longer supported',
    why: 'No security updates',
    severity: 'urgent',
    filter: { key: 'os', value: 'Unsupported' },
    count: (devices) => applyFilters(devices, { os: 'Unsupported' }).length,
  },
  {
    id: 'no-antivirus',
    title: 'No antivirus running',
    why: 'Nothing is protecting it',
    severity: 'urgent',
    filter: { key: 'av', value: 'Unprotected' },
    count: (devices) => applyFilters(devices, { av: 'Unprotected' }).length,
  },
  {
    id: 'unlicensed-office',
    title: 'Unlicensed Office',
    why: 'Not a company copy of Office',
    severity: 'warning',
    filter: { key: 'license', value: 'Unlicensed' },
    count: (devices) => applyFilters(devices, { license: 'Unlicensed' }).length,
  },
  {
    id: 'server-over-wifi',
    title: 'Server over Wi‑Fi',
    why: 'Works off the server with no cable — slow and drops out',
    severity: 'warning',
    filter: { key: 'server', value: 'Bottleneck' },
    count: (devices) => applyFilters(devices, { server: 'Bottleneck' }).length,
  },
  {
    id: 'form-factor',
    title: "Laptop or desktop doesn't suit the desk",
    why: "A suggestion, not a fault",
    severity: 'suggestion',
    filter: { key: 'formfit', value: '1' },
    count: (devices) => applyFilters(devices, { formfit: '1' }).length,
  },
  {
    id: 'stale-scan',
    title: 'Not scanned in 6 months',
    why: 'What we know about it may be out of date',
    severity: 'upkeep',
    filter: { key: 'stale', value: '1' },
    // Passes `now` through, unlike the 'stale' matcher, so a fixed clock works in tests.
    count: (devices, now) => devices.filter((device) => isStale(device, now)).length,
  },
];

/**
 * A row counted over machines IN FLEET only, and opened with the same single
 * filter: the List's own status default is In fleet, so the count and the
 * rows it opens are the same machines by construction.
 */
const fleetRow = (id, title, why, severity, key, value) => ({
  id,
  title,
  why,
  severity,
  filter: { key, value },
  count: (devices) => applyFilters(devices.filter(inFleet), { [key]: value }).length,
});

ITEMS.push(
  fleetRow('windows-not-activated', 'Windows not activated',
    'Windows will nag, lock personalisation and stop some updates.',
    'urgent', 'winact', 'Not activated'),
  fleetRow('personal-office', 'Personal Office on a work machine',
    'A Home or Personal copy, or a personal Microsoft account signed in to Office.',
    'warning', 'officewarn', 'yes'),
  fleetRow('not-joined', 'Not joined to PMW Group',
    'Not Azure AD or domain joined, so company sign-in and policies do not reach them.',
    'warning', 'joined', 'Not joined'),
  fleetRow('old-scan', 'Scanned with the old scan',
    'No licences, part numbers or software for these yet. Scan them with PMW-Device-Scan.exe.',
    'upkeep', 'scan', 'old'),
);

/**
 * The open items for a set of machines, worst first. Items with nothing to do
 * are left out, so an empty result means the fleet is in shape.
 */
export function todoItems(devices, now = Date.now()) {
  return ITEMS
    .map(({ count, ...item }) => ({ ...item, count: count(devices, now) }))
    .filter((item) => item.count > 0)
    .sort((a, b) =>
      SEVERITY_ORDER.indexOf(a.severity) - SEVERITY_ORDER.indexOf(b.severity)
      || b.count - a.count);
}

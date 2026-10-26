import { GRADES, PARTS } from './standards/defaultStandard';
import { IN_FLEET, labelOf, ramBucket, UNASSIGNED } from './deviceFilters';
import { securityRiskOf } from './health';
import { SCAN_LABELS } from './fleetFacts.js';

/**
 * What the List view's "+ Filter" menu offers, and how an active filter reads
 * on its chip. Pure: the menu is built from the machines that exist, so it
 * never offers a department or a licence nobody has.
 *
 * Every option is { filter, value, label } and choosing one writes that
 * `filter` key with that `value` -- the same keys `applyFilters` reads.
 */

export const STATUS_LABELS = { [IN_FLEET]: 'In use or in repair', All: 'Everything' };

const FLAG_LABELS = {
  work: 'Needs work',
  critical: 'Has a critical part',
  formfit: 'Form factor mismatch',
  stale: 'Not scanned in 6 months',
  officewarn: 'Personal Office',
};

const FILTER_LABELS = {
  location: 'Location', department: 'Department', status: 'Status',
  part: 'Part grade', persona: 'Workload profile', os: 'Windows', av: 'Antivirus',
  license: 'Office licence', risk: 'Security risk', type: 'Type', ram: 'RAM',
  storage: 'Storage', cpu: 'CPU age', server: 'Server link', q: 'Search',
  attention: 'Needs attention', windows: 'Windows',
  winlicence: 'Windows licence', winact: 'Windows activation', officelic: 'Office copy',
  joined: 'Joined to', software: 'Has software', scan: 'Scan',
};

/** Flags carry no value worth reading, so their chip is the label alone. */
export const isFlagFilter = (key) => key in FLAG_LABELS;

/** What the menu offers for risk: the bands, never the stored words, so 'Critical' means a part grade only. */
const RISK_BANDS = ['High', 'Medium', 'Low'];

const sortedUnique = (values) =>
  [...new Set(values)].sort((a, b) => a.localeCompare(b));

const ramSize = (label) => {
  const size = parseFloat(label);
  return Number.isFinite(size) ? size : Number.POSITIVE_INFINITY;
};

/** `sub` is the small heading the menu prints above a run of options; one-off picks have none and come first. */
const option = (filter, value, label = value, sub = FILTER_LABELS[filter] ?? null) => ({ filter, value, label, sub });
const pickOnce = (filter, value, label) => option(filter, value, label, null);

/** The label a chip shows for one active filter. */
export function chipLabel(key, value) {
  if (isFlagFilter(key)) return FLAG_LABELS[key];
  if (key === 'status') return STATUS_LABELS[value] ?? `Status: ${value}`;
  if (key === 'scan') return SCAN_LABELS[value] ?? `Scan: ${value}`;
  if (key === 'risk') return `Security risk: ${RISK_BANDS.includes(value) ? value : securityRiskOf(value)}`;
  if (key === 'os' && value === 'Unsupported') return 'Windows unsupported';
  if (key === 'av' && value === 'Unprotected') return 'No antivirus';
  if (key === 'server' && value === 'Bottleneck') return 'Server over Wi-Fi';
  if (key === 'part') {
    const [part, grade] = String(value).split(':');
    const label = PARTS.find((p) => p.key === part)?.label;
    if (label && grade) return `${label}: ${grade}`;
  }
  return `${FILTER_LABELS[key] ?? key}: ${value}`;
}


/** The menu's groups, each a title and its options, in the order the menu shows them. */
export function buildFilterGroups(devices, locationOptions = []) {
  const locations = sortedUnique([...locationOptions, UNASSIGNED]);

  const riskBands = RISK_BANDS.filter((band) => devices.some((d) => securityRiskOf(d.riskLevel) === band));

  const ramValues = sortedUnique(devices.map((d) => ramBucket(d.installedRamGB)))
    .sort((a, b) => ramSize(a) - ramSize(b));

  const parts = PARTS.flatMap((part) => GRADES.map((grade) =>
    option('part', `${part.key}:${grade}`, `${part.label} · ${grade}`)));

  return [
    {
      id: 'place',
      title: 'Place',
      items: [
        ...locations.map((l) => option('location', l)),
        ...sortedUnique(devices.map((d) => labelOf(d.department)))
          .map((d) => option('department', d)),
        option('status', IN_FLEET, STATUS_LABELS[IN_FLEET]),
        option('status', 'In use'),
        option('status', 'In repair'),
        option('status', 'Spare'),
        option('status', 'Retired'),
        option('status', 'All', 'Everything'),
      ],
    },
    {
      id: 'health',
      title: 'Health',
      items: [
        pickOnce('work', 'yes', FLAG_LABELS.work),
        pickOnce('critical', 'yes', FLAG_LABELS.critical),
        pickOnce('formfit', 'yes', FLAG_LABELS.formfit),
        ...parts,
        ...sortedUnique(devices.map((d) => labelOf(d.personaLabel)))
          .map((p) => option('persona', p)),
      ],
    },
    {
      id: 'security',
      title: 'Security',
      items: [
        pickOnce('os', 'Unsupported', 'Windows unsupported'),
        pickOnce('av', 'Unprotected', 'No antivirus'),
        ...sortedUnique(devices.map((d) => labelOf(d.licenseStatus)))
          .map((l) => option('license', l)),
        ...riskBands.map((band) => option('risk', band)),
        ...sortedUnique(devices.map((d) => labelOf(d.windowsLicenseType)))
          .map((w) => option('winlicence', w)),
        ...sortedUnique(devices.map((d) => labelOf(d.windowsActivation)))
          .map((w) => option('winact', w)),
        ...sortedUnique(devices.map((d) => labelOf(d.officeLicenseType)))
          .map((o) => option('officelic', o)),
        pickOnce('officewarn', 'yes', FLAG_LABELS.officewarn),
        ...sortedUnique(devices.map((d) => labelOf(d.deviceJoin)))
          .map((j) => option('joined', j)),
      ],
    },
    {
      id: 'hardware',
      title: 'Hardware & upkeep',
      items: [
        pickOnce('server', 'Bottleneck', 'Server over Wi-Fi'),
        pickOnce('stale', 'yes', FLAG_LABELS.stale),
        ...sortedUnique(devices.map((d) => labelOf(d.deviceType)))
          .map((t) => option('type', t)),
        ...ramValues.map((r) => option('ram', r)),
        ...sortedUnique(devices.map((d) => labelOf(d.storageType)))
          .map((s) => option('storage', s)),
        ...sortedUnique(devices.map((d) => labelOf(d.cpuAgeBand)))
          .map((c) => option('cpu', c)),
        option('scan', 'new', SCAN_LABELS.new),
        option('scan', 'basic', SCAN_LABELS.basic),
        option('scan', 'old', SCAN_LABELS.old),
      ],
    },
  ];
}

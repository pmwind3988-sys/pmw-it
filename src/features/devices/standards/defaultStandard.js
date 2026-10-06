import { PERSONAS } from '../derive/persona.js';

/**
 * What IT has decided counts as Critical, Needs attention, Moderate and
 * Optimal -- for each PART, against the work each profile does. This is the
 * first version; the one in use is the newest valid row of the
 * `IT Device Standards` list, edited on the Standards tab.
 */
export const GRADES = ['Critical', 'Needs attention', 'Moderate', 'Optimal'];
export const UNKNOWN = 'Unknown';
export const PARTS = [
  { key: 'cpu', label: 'CPU' },
  { key: 'ram', label: 'RAM' },
  { key: 'storage', label: 'Storage' },
  { key: 'graphics', label: 'Graphics' },
  { key: 'windows', label: 'Windows' },
];
export const PROFILE_KEYS = ['heavy', 'desk', 'mobile'];
export const PROFILE_SHORT = { heavy: 'Engineering', desk: 'Desk', mobile: 'Field' };
export const FALLBACK_PROFILE = 'desk';
export const STORAGE_TYPES = ['HDD only', 'Mixed', 'SSD only'];
export const GRAPHICS_KINDS = ['dedicated', 'builtIn'];
export const WINDOWS_KINDS = ['win11', 'win10Supported', 'outOfSupport', 'other'];
export const DEFAULT_COLORS = {
  Critical: '#b91c1c',
  'Needs attention': '#f59e0b',
  Moderate: '#1e3a8a',
  Optimal: '#16a34a',
  Unknown: '#8a97a8',
};

const cut = (criticalBelow, attentionBelow, optimalFrom) => ({ criticalBelow, attentionBelow, optimalFrom });

const profile = (persona, { cpu, ram, storageSize, builtIn }) => ({
  label: persona.label,
  blurb: persona.blurb,
  prefers: persona.prefers,
  cpu,
  ram,
  storageSize,
  storageType: { 'HDD only': 'Critical', Mixed: 'Needs attention', 'SSD only': 'Optimal' },
  graphics: { dedicated: 'Optimal', builtIn },
  windows: { win11: 'Optimal', win10Supported: 'Moderate', outOfSupport: 'Critical', other: 'Critical' },
});

export function defaultStandard() {
  return {
    schema: 1,
    profiles: {
      heavy: profile(PERSONAS.HEAVY, { cpu: cut(8, 10, 12), ram: cut(8, 16, 32), storageSize: cut(128, 256, 512), builtIn: 'Needs attention' }),
      desk: profile(PERSONAS.DESK, { cpu: cut(6, 8, 10), ram: cut(8, 8, 16), storageSize: cut(128, 256, 256), builtIn: 'Optimal' }),
      mobile: profile(PERSONAS.MOBILE, { cpu: cut(7, 10, 12), ram: cut(8, 16, 16), storageSize: cut(128, 256, 512), builtIn: 'Optimal' }),
    },
    departments: {
      ENGINEERING: 'heavy', PRODUCTION: 'heavy', QAQC: 'heavy', QC: 'heavy', IT: 'heavy', MARKETING: 'heavy',
      SALES: 'mobile', ADMIN: 'mobile',
      LOGISTICS: 'desk', SHIPPING: 'desk', PURCHASING: 'desk', STORE: 'desk', STOCKYARD: 'desk',
      STOCKYARDF1: 'desk', GUARDHOUSE: 'desk', 'PML GUARDHOUSE': 'desk', FINANCE: 'desk', ACCOUNT: 'desk', HR: 'desk',
    },
    colors: { ...DEFAULT_COLORS },
  };
}

export const DEFAULT_STANDARD = defaultStandard();

export const cloneStandard = (standard) => JSON.parse(JSON.stringify(standard));

/** The profile a department is judged against; anything unmapped is Desk, as before. */
export function profileKeyFor(standard, department) {
  const key = String(department ?? '').trim().toUpperCase();
  const profileKey = key ? standard.departments?.[key] : null;
  return PROFILE_KEYS.includes(profileKey) ? profileKey : FALLBACK_PROFILE;
}

# Device Standards Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** IT sets, on a Standards tab, what counts as Critical / Needs attention / Moderate / Optimal for each part (CPU, RAM, Storage, Graphics, Windows) of each workload profile, plus the grade colours. Every machine is then graded per part with no overall grade.

**Architecture:**
- **The standard is data.** It is a JSON document stored append-only in a SharePoint list, `IT Device Standards`. It is validated and diffed by pure modules in `src/features/devices/standards/`.
- **Grading is a pure function.** `derive/partGrades.js` is applied on read with the loaded standard, through `useDevices`, so a saved standard regrades the fleet with no re-scan.
- **The old fleet-wide verdict goes.** `deviceFit.js` and its `fitStatus` / `fitReasons` are removed. Every screen that used them reads the per-part grades instead.
- **Colours are CSS variables** (`--grade-*`) set on the device section's root element.

**Tech Stack:** React 19, React Router, Vite 8, Vitest 3, SharePoint REST via `spFetch`.

**Spec:** `docs/superpowers/specs/2026-10-06-device-standards-design.md`. Read it before starting any task. The design mockup is screen 8 of https://claude.ai/artifact/PuQ6YuUwvwWG8aumwkvBjD.

## Global Constraints

- **Where to work.**
  - Use the worktree `C:\Users\User\pmw-it\.claude\worktrees\devices-map`, branch `devices-map-lifecycle`.
  - Never touch `C:\Users\User\pmw-it` itself.
  - Do not modify `src/App.jsx` or any checklist-link file (`server/`, `api/`, `src/public/`, `src/features/forms/links/`).
- **Grades.**
  - The grade names are exactly `Critical`, `Needs attention`, `Moderate`, `Optimal`, plus `Unknown`. The new code uses lower-case "attention". The old `'Needs Attention'` spelling disappears with `deviceFit.js`.
  - The parts are exactly `cpu`, `ram`, `storage`, `graphics`, `windows`, labelled `CPU`, `RAM`, `Storage`, `Graphics`, `Windows`.
  - The profile keys are exactly `heavy`, `desk`, `mobile`, short names `Engineering`, `Desk`, `Field`. An unmapped department, or no department, uses `desk`.
- **Cut-offs.**
  - A value `< criticalBelow` is Critical, `< attentionBelow` is Needs attention, `< optimalFrom` is Moderate, anything else is Optimal.
  - A valid standard has `criticalBelow ≤ attentionBelow ≤ optimalFrom`. Equal values are allowed: the band between them is empty.
  - Storage takes the WORSE of the disk-type grade and the size grade.
- **The list.**
  - It is titled exactly `IT Device Standards`, with columns `Standard` (note, `RichText: false`), `SavedOn` (datetime, `DisplayFormat: 1`), `SavedBy` (text), `Summary` (note).
  - It is append-only: never update or delete a row.
  - Create it only through `provisionSchema`.
- **Colours.**
  - The device section never hard-codes a grade colour. It uses `var(--grade-critical|attention|moderate|optimal|unknown)` and the matching `-ink` text colour.
  - Default colours are `#dc2626`, `#f59e0b`, `#1a88de`, `#12a150`, `#8a97a8`.
- **Not touched:** `riskScore.js`, `officeLicense.js`, `gpuClass.js`, `serverDependency.js`.
- **React rules** (AGENTS.md):
  - No `navigate()` in effects.
  - No state reset from an effect.
  - No ref writes during render.
  - No helper exported from a component file.
  - Import every icon used.
  - Inputs at least 16px, touch targets at least 44px.
  - `<Button loading>` on any button that starts a SharePoint write.
- **No literal invisible characters** in source. Visible symbols like `≤ → · …` are fine.
- **Checks.**
  - `npx eslint src` may report only the two pre-existing errors, in `ThemeContext.jsx` and `SemanticContext.jsx`.
  - Tests: `npx vitest run <path>` for one area, `npm test` for everything.
- **Commits** end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.

## File map

| File | Status | Responsibility |
|---|---|---|
| `src/features/devices/standards/defaultStandard.js` | create | grades, parts, profiles, the default standard, `profileKeyFor`, `cloneStandard` |
| `src/features/devices/standards/validateStandard.js` | create | is a standard usable |
| `src/features/devices/standards/gradeColors.js` | create | ink colour, contrast, CSS variables, grade slugs |
| `src/features/devices/standards/diffStandard.js` | create | the change summary |
| `src/features/devices/standards/previewChanges.js` | create | which machines would change grade |
| `src/features/devices/derive/partGrades.js` | create | device + standard → per-part grades |
| `src/features/devices/derive/regrade.js` | create | apply a standard to a list of devices |
| `src/features/devices/derive/deviceFit.js` (+ test) | delete | replaced by partGrades |
| `src/features/devices/derive/enrichFit.js` | modify | grade with the default standard on read |
| `src/features/devices/derive/persona.js` | modify | keep the profile labels; drop `personaFor` / `BY_DEPARTMENT` (moved into the standard) |
| `src/features/devices/sharepoint/standardsSchema.js` | create | the list and its rows |
| `src/features/devices/sharepoint/readStandards.js` | create | the newest valid version + history; may I edit |
| `src/features/devices/sharepoint/saveStandard.js` | create | append a version |
| `src/features/devices/useStandards.js` | create | loads the standard for a page |
| `src/features/devices/useDevices.js` | modify | grades rows with the loaded standard; returns `standards` |
| `src/features/devices/map/zones.js` | modify | per-part tile summaries; sort by critical parts |
| `src/features/devices/ui/PartBars.jsx`, `PartChips.jsx`, `PartTable.jsx`, `StandardsPage.jsx` | create | the new UI |
| `src/features/devices/ui/MapTile.jsx`, `MachineCard.jsx`, `DepartmentView.jsx`, `LocationView.jsx`, `DeviceMap.jsx`, `WorldMap.jsx` | modify | use the part grades |
| `src/features/devices/ui/DeviceTable.jsx`, `DeviceCharts.jsx`, `DepartmentHeatmap.jsx` | modify | part columns, critical-parts chart, part heatmap |
| `src/features/devices/stats/deviceStats.js`, `deviceFilters.js`, `multiValue.js` | modify | part-based figures and filters |
| `src/pages/DevicesPage.jsx`, `src/pages/DeviceDetailPage.jsx` | modify | the Standards tab, grade colours, the part table |
| `src/styles/devices.css` | modify | `.g-*` grade classes, part bars and chips, the Standards page |
| `AGENTS.md` | modify | conventions and where-to-look |

---

### Task 1: The standard itself, and what makes one valid

**Files:**
- Create: `src/features/devices/standards/defaultStandard.js`, `src/features/devices/standards/defaultStandard.test.js`
- Create: `src/features/devices/standards/validateStandard.js`, `src/features/devices/standards/validateStandard.test.js`

**Interfaces:**
- Consumes: `PERSONAS` from `src/features/devices/derive/persona.js`. Each persona has `label`, `blurb` and `prefers`.
- Produces:
  - `GRADES`, `UNKNOWN`, `PARTS` (`[{ key, label }]`), `PROFILE_KEYS`, `PROFILE_SHORT`, `FALLBACK_PROFILE`, `STORAGE_TYPES`, `GRAPHICS_KINDS`, `WINDOWS_KINDS`, `DEFAULT_COLORS`
  - `defaultStandard()` → a fresh standard
  - `DEFAULT_STANDARD`
  - `cloneStandard(s)`
  - `profileKeyFor(standard, department)` → `'heavy'|'desk'|'mobile'`
  - `validateStandard(standard)` → `{ ok, errors: [{ path, message }] }`

A standard looks like this:

```
{ schema: 1,
  profiles: { heavy|desk|mobile: {
    label, blurb, prefers,
    cpu: { criticalBelow, attentionBelow, optimalFrom },
    ram: { … }, storageSize: { … },
    storageType: { 'HDD only': grade, Mixed: grade, 'SSD only': grade },
    graphics: { dedicated: grade, builtIn: grade },
    windows: { win11, win10Supported, outOfSupport, other } } },
  departments: { ENGINEERING: 'heavy', … },
  colors: { Critical, 'Needs attention', Moderate, Optimal, Unknown } }
```

- [ ] **Step 1: Write the failing tests**

Create `src/features/devices/standards/defaultStandard.test.js`:

```js
import { describe, it, expect } from 'vitest';
import {
  DEFAULT_STANDARD, defaultStandard, cloneStandard, profileKeyFor, GRADES, PARTS, PROFILE_KEYS,
} from './defaultStandard.js';
import { validateStandard } from './validateStandard.js';

describe('the default standard', () => {
  it('is valid', () => {
    expect(validateStandard(DEFAULT_STANDARD)).toEqual({ ok: true, errors: [] });
  });

  it('keeps today\u2019s RAM floors and comfort levels', () => {
    expect(DEFAULT_STANDARD.profiles.heavy.ram).toEqual({ criticalBelow: 8, attentionBelow: 16, optimalFrom: 32 });
    expect(DEFAULT_STANDARD.profiles.desk.ram).toEqual({ criticalBelow: 8, attentionBelow: 8, optimalFrom: 16 });
    expect(DEFAULT_STANDARD.profiles.mobile.ram).toEqual({ criticalBelow: 8, attentionBelow: 16, optimalFrom: 16 });
  });

  it('asks a graphics card of engineering only', () => {
    expect(DEFAULT_STANDARD.profiles.heavy.graphics.builtIn).toBe('Needs attention');
    expect(DEFAULT_STANDARD.profiles.desk.graphics.builtIn).toBe('Optimal');
  });

  it('names exactly four grades and five parts', () => {
    expect(GRADES).toEqual(['Critical', 'Needs attention', 'Moderate', 'Optimal']);
    expect(PARTS.map((p) => p.key)).toEqual(['cpu', 'ram', 'storage', 'graphics', 'windows']);
    expect(PROFILE_KEYS).toEqual(['heavy', 'desk', 'mobile']);
  });

  it('hands out a fresh copy every time', () => {
    const a = defaultStandard();
    a.profiles.heavy.ram.optimalFrom = 99;
    expect(defaultStandard().profiles.heavy.ram.optimalFrom).toBe(32);
    const b = cloneStandard(DEFAULT_STANDARD);
    b.colors.Critical = '#000000';
    expect(DEFAULT_STANDARD.colors.Critical).toBe('#dc2626');
  });
});

describe('profileKeyFor', () => {
  it('maps a known department, whatever its case or spacing', () => {
    expect(profileKeyFor(DEFAULT_STANDARD, ' engineering ')).toBe('heavy');
    expect(profileKeyFor(DEFAULT_STANDARD, 'SALES')).toBe('mobile');
  });

  it('falls back to Desk for an unknown or missing department', () => {
    expect(profileKeyFor(DEFAULT_STANDARD, 'WAREHOUSE 9')).toBe('desk');
    expect(profileKeyFor(DEFAULT_STANDARD, null)).toBe('desk');
  });
});
```

Create `src/features/devices/standards/validateStandard.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { validateStandard } from './validateStandard.js';
import { defaultStandard } from './defaultStandard.js';

const broken = (edit) => { const s = defaultStandard(); edit(s); return validateStandard(s); };

describe('validateStandard', () => {
  it('refuses cut-offs out of order, naming where', () => {
    const result = broken((s) => { s.profiles.heavy.ram.criticalBelow = 40; });
    expect(result.ok).toBe(false);
    expect(result.errors[0]).toEqual({ path: 'profiles.heavy.ram', message: 'Must run in order: Critical below ≤ Needs attention below ≤ Optimal from.' });
  });

  it('accepts equal cut-offs: an empty band', () => {
    expect(broken((s) => { s.profiles.desk.cpu = { criticalBelow: 8, attentionBelow: 8, optimalFrom: 8 }; }).ok).toBe(true);
  });

  it('refuses a missing or negative number', () => {
    expect(broken((s) => { s.profiles.desk.cpu.optimalFrom = null; }).ok).toBe(false);
    expect(broken((s) => { s.profiles.desk.cpu.criticalBelow = -1; }).ok).toBe(false);
  });

  it('refuses a choice that is not a grade', () => {
    expect(broken((s) => { s.profiles.mobile.windows.win11 = 'Great'; }).errors[0].path).toBe('profiles.mobile.windows.win11');
  });

  it('refuses a missing profile, an unknown profile on a department, and a bad colour', () => {
    expect(broken((s) => { delete s.profiles.desk; }).ok).toBe(false);
    expect(broken((s) => { s.departments.HR = 'boss'; }).errors[0].path).toBe('departments.HR');
    expect(broken((s) => { s.colors.Optimal = 'green'; }).errors[0].path).toBe('colors.Optimal');
  });

  it('refuses a standard from a newer schema, and nothing at all', () => {
    expect(broken((s) => { s.schema = 2; }).ok).toBe(false);
    expect(validateStandard(null).ok).toBe(false);
  });
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `npx vitest run src/features/devices/standards`
Expected: FAIL. The modules don't exist.

- [ ] **Step 3: Implement**

Create `src/features/devices/standards/defaultStandard.js`:

```js
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
  Critical: '#dc2626',
  'Needs attention': '#f59e0b',
  Moderate: '#1a88de',
  Optimal: '#12a150',
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
```

Create `src/features/devices/standards/validateStandard.js`:

```js
import {
  GRADES, UNKNOWN, PROFILE_KEYS, STORAGE_TYPES, GRAPHICS_KINDS, WINDOWS_KINDS,
} from './defaultStandard.js';

const CUT_PARTS = ['cpu', 'ram', 'storageSize'];
const ORDER = 'Must run in order: Critical below ≤ Needs attention below ≤ Optimal from.';

/** Whether a standard can be graded with, and every reason it cannot, each with where. */
export function validateStandard(standard) {
  const errors = [];
  const add = (path, message) => errors.push({ path, message });

  if (!standard || typeof standard !== 'object') {
    return { ok: false, errors: [{ path: '', message: 'The standard is empty.' }] };
  }
  if (standard.schema !== 1) add('schema', 'This standard was saved by a newer version of the portal.');

  for (const key of PROFILE_KEYS) {
    const profile = standard.profiles?.[key];
    if (!profile) {
      add(`profiles.${key}`, 'This profile is missing.');
      continue;
    }
    for (const part of CUT_PARTS) {
      const c = profile[part];
      const nums = [c?.criticalBelow, c?.attentionBelow, c?.optimalFrom];
      if (!nums.every((n) => typeof n === 'number' && Number.isFinite(n) && n >= 0)) {
        add(`profiles.${key}.${part}`, 'Every cut-off needs a number of 0 or more.');
      } else if (!(nums[0] <= nums[1] && nums[1] <= nums[2])) {
        add(`profiles.${key}.${part}`, ORDER);
      }
    }
    const choices = [['storageType', STORAGE_TYPES], ['graphics', GRAPHICS_KINDS], ['windows', WINDOWS_KINDS]];
    for (const [part, kinds] of choices) {
      for (const kind of kinds) {
        if (!GRADES.includes(profile[part]?.[kind])) add(`profiles.${key}.${part}.${kind}`, 'Pick a grade.');
      }
    }
  }

  for (const [department, profileKey] of Object.entries(standard.departments ?? {})) {
    if (!PROFILE_KEYS.includes(profileKey)) add(`departments.${department}`, 'Unknown profile.');
  }

  for (const grade of [...GRADES, UNKNOWN]) {
    if (!/^#[0-9a-f]{6}$/i.test(standard.colors?.[grade] ?? '')) add(`colors.${grade}`, 'Use a colour like #1a88de.');
  }

  return { ok: errors.length === 0, errors };
}
```

- [ ] **Step 4: Run them to see them pass**

Run: `npx vitest run src/features/devices/standards`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/standards
git commit -m "Write the device standard down as data: per-profile cut-offs and grades for each part, and what makes one valid

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Grade each part of a machine

**Files:**
- Create: `src/features/devices/derive/partGrades.js`, `src/features/devices/derive/partGrades.test.js`
- Create: `src/features/devices/derive/regrade.js`, `src/features/devices/derive/regrade.test.js`
- Modify: `src/features/devices/derive/enrichFit.js`, `src/features/devices/derive/enrichFit.test.js`, `src/features/devices/multiValue.js`
- Delete: `src/features/devices/derive/deviceFit.js`, `src/features/devices/derive/deviceFit.test.js`

**Interfaces:**
- Consumes: Task 1 exports.
- Produces:
  - `cutoffGrade(value, cut)` → grade or `null`
  - `partGrades(device, standard)`, which returns:
    - `personaKey`, `personaLabel`, `personaBlurb`
    - `suggestedFormFactor`, `formFactorNote`, `formFactorMatches`
    - `parts`: `{ cpu|ram|storage|graphics|windows: { grade, value, reason } }`
    - `gradeCpu`, `gradeRam`, `gradeStorage`, `gradeGraphics`, `gradeWindows`
    - `criticalParts` and `attentionParts`: arrays of part keys in `PARTS` order
    - `actionRequired`
  - `regrade(devices, standard)` → the devices with `partGrades` laid over each one
- Removed: `deviceFit`, `FIT_LEVELS`, and the device fields `fitStatus` and `fitReasons`.

- [ ] **Step 1: Write the failing tests**

Create `src/features/devices/derive/partGrades.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { partGrades, cutoffGrade } from './partGrades.js';
import { defaultStandard } from '../standards/defaultStandard.js';

const S = defaultStandard();
const machine = (over) => ({
  department: 'ENGINEERING', scanComplete: true, deviceType: 'Laptop',
  cpuGenerationRank: 12, cpuAgeBand: 'Current', cpuModel: 'i7-1265U',
  installedRamGB: 32, storageTotalGB: 512, storageType: 'SSD only',
  dedicatedGpu: true, osSupported: true, windowsMajor: 11, windowsVersion: 'Microsoft Windows 11 Pro',
  ...over,
});

describe('cutoffGrade', () => {
  const c = { criticalBelow: 8, attentionBelow: 16, optimalFrom: 32 };
  it.each([[7, 'Critical'], [8, 'Needs attention'], [15, 'Needs attention'], [16, 'Moderate'], [31, 'Moderate'], [32, 'Optimal']])(
    '%i GB is %s', (value, grade) => expect(cutoffGrade(value, c)).toBe(grade),
  );
  it('skips an empty band', () => {
    expect(cutoffGrade(8, { criticalBelow: 8, attentionBelow: 8, optimalFrom: 16 })).toBe('Moderate');
  });
  it('has no grade for a missing value', () => expect(cutoffGrade(null, c)).toBeNull());
});

describe('partGrades', () => {
  it('grades a well-specified engineering machine Optimal on every part', () => {
    const g = partGrades(machine(), S);
    expect([g.gradeCpu, g.gradeRam, g.gradeStorage, g.gradeGraphics, g.gradeWindows]).toEqual(Array(5).fill('Optimal'));
    expect(g.actionRequired).toBe('Nothing to do');
    expect(g.personaKey).toBe('heavy');
  });

  it('grades each part on its own, with a reason naming the threshold', () => {
    const g = partGrades(machine({ installedRamGB: 16 }), S);
    expect(g.parts.ram).toEqual({ grade: 'Moderate', value: '16 GB', reason: 'Under the 32 GB needed for Optimal' });
    expect(g.gradeCpu).toBe('Optimal');
  });

  it('judges the same machine differently on a different desk', () => {
    expect(partGrades(machine({ installedRamGB: 16, department: 'FINANCE' }), S).gradeRam).toBe('Optimal');
  });

  it('makes an obsolete processor family Critical whatever the cut-offs', () => {
    const g = partGrades(machine({ cpuAgeBand: 'Obsolete', cpuGenerationRank: null, cpuModel: 'Celeron N4020' }), S);
    expect(g.parts.cpu.grade).toBe('Critical');
  });

  it('takes the worse of disk type and disk size', () => {
    expect(partGrades(machine({ storageTotalGB: 1000, storageType: 'HDD only' }), S).gradeStorage).toBe('Critical');
    expect(partGrades(machine({ storageTotalGB: 200, storageType: 'SSD only' }), S).gradeStorage).toBe('Needs attention');
  });

  it('asks a graphics card of engineering only', () => {
    expect(partGrades(machine({ dedicatedGpu: false }), S).gradeGraphics).toBe('Needs attention');
    expect(partGrades(machine({ dedicatedGpu: false, department: 'HR' }), S).gradeGraphics).toBe('Optimal');
  });

  it('places every Windows version', () => {
    const win10 = { windowsMajor: 10, windowsVersion: 'Microsoft Windows 10 Pro' };
    expect(partGrades(machine(win10), S).gradeWindows).toBe('Moderate');
    expect(partGrades(machine({ ...win10, osSupported: false }), S).gradeWindows).toBe('Critical');
    expect(partGrades(machine({ windowsMajor: 7, windowsVersion: 'Windows 7 Pro', osSupported: null }), S).gradeWindows).toBe('Critical');
    expect(partGrades(machine({ windowsMajor: null, windowsVersion: null, osSupported: null }), S).gradeWindows).toBe('Unknown');
  });

  it('never guesses a value the scan did not report', () => {
    const g = partGrades(machine({ cpuGenerationRank: null, installedRamGB: null, dedicatedGpu: null }), S);
    expect([g.gradeCpu, g.gradeRam, g.gradeGraphics]).toEqual(['Unknown', 'Unknown', 'Unknown']);
  });

  it('names the parts to act on', () => {
    const g = partGrades(machine({ installedRamGB: 4, storageType: 'HDD only', cpuGenerationRank: 9 }), S);
    expect(g.criticalParts).toEqual(['ram', 'storage']);
    expect(g.attentionParts).toEqual(['cpu']);
    expect(g.actionRequired).toBe('Upgrade now: RAM, Storage');
    expect(partGrades(machine({ cpuGenerationRank: 9 }), S).actionRequired).toBe('Plan: CPU');
  });

  it('grades nothing on an incomplete scan', () => {
    const g = partGrades(machine({ scanComplete: false }), S);
    expect(g.gradeRam).toBe('Unknown');
    expect(g.actionRequired).toBe('Re-run the scan');
  });

  it('uses the department map the standard carries', () => {
    const custom = defaultStandard();
    custom.departments.FINANCE = 'heavy';
    expect(partGrades(machine({ department: 'FINANCE', installedRamGB: 16 }), custom).gradeRam).toBe('Moderate');
  });
});
```

Create `src/features/devices/derive/regrade.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { regrade } from './regrade.js';
import { defaultStandard } from '../standards/defaultStandard.js';

describe('regrade', () => {
  it('lays the grades over every device without losing its own fields', () => {
    const [d] = regrade([{ id: 3, computerName: 'PC1', department: 'HR', installedRamGB: 4, scanComplete: true }], defaultStandard());
    expect(d).toMatchObject({ id: 3, computerName: 'PC1', gradeRam: 'Critical' });
  });

  it('changes a grade when the standard changes', () => {
    const strict = defaultStandard();
    strict.profiles.desk.ram = { criticalBelow: 16, attentionBelow: 16, optimalFrom: 32 };
    expect(regrade([{ department: 'HR', installedRamGB: 8, scanComplete: true }], strict)[0].gradeRam).toBe('Critical');
  });
});
```

In `src/features/devices/derive/enrichFit.test.js`, replace every assertion on `fitStatus` / `fitReasons` / `personaKey` with assertions on the new shape. Keep the test names' intent: the persona layer is laid over the record. For example, `expect(result.gradeRam).toBe(…)` and `expect(result.personaKey).toBe('heavy')`. Read the file first; keep its fixtures.

- [ ] **Step 2: Run them to see them fail**

Run: `npx vitest run src/features/devices/derive`
Expected: FAIL. `partGrades` and `regrade` don't exist.

- [ ] **Step 3: Implement**

Create `src/features/devices/derive/partGrades.js`:

```js
import { GRADES, UNKNOWN, PARTS, profileKeyFor } from '../standards/defaultStandard.js';

/**
 * Every part of a machine graded on its own, against the standard for the
 * desk it sits on. There is deliberately NO overall grade: a fast processor
 * does not make up for a hard disk, and one word for the whole machine hid
 * which part to fix.
 */
const RANK = Object.fromEntries(GRADES.map((grade, index) => [grade, index]));
const worse = (a, b) => {
  if (!a) return b;
  if (!b) return a;
  return RANK[a] <= RANK[b] ? a : b;
};

export function cutoffGrade(value, cut) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return null;
  if (value < cut.criticalBelow) return 'Critical';
  if (value < cut.attentionBelow) return 'Needs attention';
  if (value < cut.optimalFrom) return 'Moderate';
  return 'Optimal';
}

const ordinal = (n) => {
  const tens = n % 100;
  if (tens >= 11 && tens <= 13) return `${n}th`;
  return `${n}${({ 1: 'st', 2: 'nd', 3: 'rd' })[n % 10] ?? 'th'}`;
};
const gen = (n) => `${ordinal(n)} gen`;
const gb = (n) => `${n} GB`;

function reasonFor(grade, cut, unit) {
  if (grade === 'Critical') return `Below the ${unit(cut.criticalBelow)} floor`;
  if (grade === 'Needs attention') return `Under the ${unit(cut.attentionBelow)} needed for Moderate`;
  if (grade === 'Moderate') return `Under the ${unit(cut.optimalFrom)} needed for Optimal`;
  return `At or above the ${unit(cut.optimalFrom)} for Optimal`;
}

const unknown = (value, reason) => ({ grade: UNKNOWN, value, reason });

function gradeCpu(device, profile) {
  if (device.cpuAgeBand === 'Obsolete') {
    return { grade: 'Critical', value: device.cpuModel ?? 'Obsolete family', reason: 'Pentium, Celeron or AMD from before Ryzen — always Critical' };
  }
  const rank = device.cpuGenerationRank;
  const grade = cutoffGrade(rank, profile.cpu);
  if (!grade) return unknown(device.cpuModel ?? '—', 'The scan did not report a processor generation');
  return { grade, value: gen(rank), reason: reasonFor(grade, profile.cpu, gen) };
}

function gradeRam(device, profile) {
  const grade = cutoffGrade(device.installedRamGB, profile.ram);
  if (!grade) return unknown('—', 'The scan did not report the memory fitted');
  return { grade, value: gb(device.installedRamGB), reason: reasonFor(grade, profile.ram, gb) };
}

function gradeStorage(device, profile) {
  const typeGrade = profile.storageType[device.storageType] ?? null;
  const sizeGrade = cutoffGrade(device.storageTotalGB, profile.storageSize);
  const grade = worse(typeGrade, sizeGrade);
  if (!grade) return unknown('—', 'The scan did not report the disks');
  const value = [
    typeof device.storageTotalGB === 'number' ? gb(device.storageTotalGB) : null,
    typeGrade ? device.storageType : null,
  ].filter(Boolean).join(' · ');
  const typeDecides = typeGrade === grade && (!sizeGrade || RANK[typeGrade] <= RANK[sizeGrade]);
  const reason = typeDecides
    ? `${device.storageType} counts as ${grade} for this work`
    : reasonFor(sizeGrade, profile.storageSize, gb);
  return { grade, value, reason };
}

function gradeGraphics(device, profile) {
  if (device.dedicatedGpu === true) {
    const grade = profile.graphics.dedicated;
    return { grade, value: 'Dedicated card', reason: `A dedicated card counts as ${grade} for this work` };
  }
  if (device.dedicatedGpu === false) {
    const grade = profile.graphics.builtIn;
    return { grade, value: 'Built-in graphics', reason: `Built-in graphics count as ${grade} for this work` };
  }
  return unknown('—', 'The scan did not report the graphics adapter');
}

const WINDOWS_LABEL = {
  win11: 'Windows 11',
  win10Supported: 'Windows 10, still supported',
  outOfSupport: 'Out of support',
  other: 'An older or unrecognised Windows',
};

function windowsKind(device) {
  const text = device.windowsVersion ?? '';
  if (device.osSupported === false) return 'outOfSupport';
  if (device.windowsMajor === 11 || /windows 11/i.test(text)) return 'win11';
  if (device.windowsMajor === 10 || /windows 10/i.test(text)) return 'win10Supported';
  if (text || typeof device.windowsMajor === 'number') return 'other';
  return null;
}

function gradeWindows(device, profile) {
  const kind = windowsKind(device);
  if (!kind) return unknown('—', 'The scan did not report the Windows version');
  const grade = profile.windows[kind];
  return { grade, value: device.windowsVersion ?? WINDOWS_LABEL[kind], reason: `${WINDOWS_LABEL[kind]} counts as ${grade}` };
}

/** The portability tag, moved unchanged from the old deviceFit: a label, never a fault. */
function portability(device, profile) {
  if (!profile.prefers) {
    return { suggestedFormFactor: null, formFactorNote: 'No department on the record, so no form factor is suggested', formFactorMatches: null };
  }
  const matches = device.deviceType === 'Unknown' || !device.deviceType ? null : device.deviceType === profile.prefers;
  const note = profile.prefers === 'Laptop'
    ? 'This role works away from the desk — a laptop suits it better'
    : 'Deskbound work with headroom to buy — a desktop gives more for the money';
  return {
    suggestedFormFactor: profile.prefers,
    formFactorNote: matches ? `Already a ${profile.prefers.toLowerCase()} — a good match for this role` : note,
    formFactorMatches: matches,
  };
}

const labels = (keys) => keys.map((key) => PARTS.find((part) => part.key === key).label).join(', ');

export function partGrades(device, standard) {
  const personaKey = profileKeyFor(standard, device.department);
  const profile = standard.profiles[personaKey];
  const incomplete = device.scanComplete === false;

  const parts = incomplete
    ? Object.fromEntries(PARTS.map(({ key }) => [key, unknown('—', 'Scan incomplete — nothing to judge')]))
    : {
      cpu: gradeCpu(device, profile),
      ram: gradeRam(device, profile),
      storage: gradeStorage(device, profile),
      graphics: gradeGraphics(device, profile),
      windows: gradeWindows(device, profile),
    };

  const criticalParts = PARTS.filter(({ key }) => parts[key].grade === 'Critical').map(({ key }) => key);
  const attentionParts = PARTS.filter(({ key }) => parts[key].grade === 'Needs attention').map(({ key }) => key);

  let actionRequired = 'Nothing to do';
  if (incomplete) actionRequired = 'Re-run the scan';
  else if (criticalParts.length) actionRequired = `Upgrade now: ${labels(criticalParts)}`;
  else if (attentionParts.length) actionRequired = `Plan: ${labels(attentionParts)}`;

  return {
    personaKey,
    personaLabel: profile.label,
    personaBlurb: profile.blurb,
    ...portability(device, profile),
    parts,
    gradeCpu: parts.cpu.grade,
    gradeRam: parts.ram.grade,
    gradeStorage: parts.storage.grade,
    gradeGraphics: parts.graphics.grade,
    gradeWindows: parts.windows.grade,
    criticalParts,
    attentionParts,
    actionRequired,
  };
}
```

Create `src/features/devices/derive/regrade.js`:

```js
import { partGrades } from './partGrades.js';

/** The fleet graded against one standard. Pure: the rows are not changed. */
export function regrade(devices, standard) {
  return devices.map((device) => ({ ...device, ...partGrades(device, standard) }));
}
```

In `src/features/devices/derive/enrichFit.js`:
- Replace `import { deviceFit } from './deviceFit.js';` with:

```js
import { partGrades } from './partGrades.js';
import { DEFAULT_STANDARD } from '../standards/defaultStandard.js';
```

- Replace the last line with:

```js
  // Graded against the DEFAULT standard here, so a record is never ungraded;
  // useDevices regrades it with the saved standard as soon as that is loaded.
  return { ...withFacts, ...partGrades(withFacts, DEFAULT_STANDARD) };
```

Delete `src/features/devices/derive/deviceFit.js` and `src/features/devices/derive/deviceFit.test.js`.

In `src/features/devices/multiValue.js`, remove `'fitReasons',` from the list.

- [ ] **Step 4: Run the device suite**

Run: `npx vitest run src/features/devices`
Expected: the new tests PASS. Tests in `zones.test.js`, `deviceStats.test.js` and `deviceFilters.test.js` that hand-build devices with `fitStatus` still pass, because they never read `deviceFit`. Tasks 6 and 8 rewrite them.

If any other test imports `deviceFit`, fix that import:

```bash
grep -rn "deviceFit" src
```

- [ ] **Step 5: Commit**

```bash
git add -A src/features/devices/derive src/features/devices/multiValue.js
git commit -m "Grade every part of a machine on its own against the standard, and drop the single machine verdict

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: Colours, change summary and preview

**Files:**
- Create: `src/features/devices/standards/gradeColors.js` (+ `.test.js`), `diffStandard.js` (+ `.test.js`), `previewChanges.js` (+ `.test.js`)

**Interfaces:**
- Consumes: Task 1 exports; `partGrades` (Task 2).
- Produces:
  - `GRADE_SLUG` (`{ Critical: 'critical', 'Needs attention': 'attention', Moderate: 'moderate', Optimal: 'optimal', Unknown: 'unknown' }`)
  - `contrastRatio(a, b)`, `inkFor(hex)` (`'#ffffff'` or `'#101828'`), `tooSimilar(colors)`, `gradeCssVars(colors)`
  - `diffStandard(before, after)` → `string[]`, and `summaryOf(lines)` → `string`
  - `previewChanges(devices, before, after)` → `[{ part, from, to, count }]`

- [ ] **Step 1: Write the failing tests**

`src/features/devices/standards/gradeColors.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { inkFor, tooSimilar, gradeCssVars, contrastRatio } from './gradeColors.js';
import { DEFAULT_COLORS } from './defaultStandard.js';

describe('gradeColors', () => {
  it('puts dark text on light colours and white text on dark ones', () => {
    expect(inkFor('#ffffff')).toBe('#101828');
    expect(inkFor('#101828')).toBe('#ffffff');
    expect(inkFor('#f59e0b')).toBe('#101828');
  });

  it('measures contrast like WCAG', () => {
    expect(contrastRatio('#000000', '#ffffff')).toBeCloseTo(21, 0);
    expect(contrastRatio('#123456', '#123456')).toBeCloseTo(1, 5);
  });

  it('warns only when two grades are too alike', () => {
    expect(tooSimilar(DEFAULT_COLORS)).toBe(false);
    expect(tooSimilar({ ...DEFAULT_COLORS, Moderate: '#12a150' })).toBe(true);
  });

  it('turns the colours into CSS variables with a text colour each', () => {
    const vars = gradeCssVars(DEFAULT_COLORS);
    expect(vars['--grade-critical']).toBe('#dc2626');
    expect(vars['--grade-attention-ink']).toBe('#101828');
    expect(Object.keys(vars)).toHaveLength(10);
  });
});
```

`src/features/devices/standards/diffStandard.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { diffStandard, summaryOf } from './diffStandard.js';
import { defaultStandard } from './defaultStandard.js';

describe('diffStandard', () => {
  it('says nothing when nothing changed', () => {
    expect(summaryOf(diffStandard(defaultStandard(), defaultStandard()))).toBe('No changes');
  });

  it('reads like a sentence for every kind of change', () => {
    const after = defaultStandard();
    after.profiles.heavy.ram.optimalFrom = 24;
    after.profiles.desk.windows.win10Supported = 'Optimal';
    after.departments.QAQC = 'desk';
    after.colors.Critical = '#b91c1c';
    expect(diffStandard(defaultStandard(), after)).toEqual([
      'Engineering RAM optimal from 32 → 24',
      'Desk Windows 10 (supported) Moderate → Optimal',
      'QAQC judged as Desk',
      'Critical colour changed',
    ]);
  });
});
```

`src/features/devices/standards/previewChanges.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { previewChanges } from './previewChanges.js';
import { defaultStandard } from './defaultStandard.js';

const device = (ram) => ({
  department: 'ENGINEERING', scanComplete: true, installedRamGB: ram, cpuGenerationRank: 12,
  storageTotalGB: 512, storageType: 'SSD only', dedicatedGpu: true, osSupported: true, windowsMajor: 11,
});

describe('previewChanges', () => {
  it('counts the machines that would change grade, per part', () => {
    const after = defaultStandard();
    after.profiles.heavy.ram.optimalFrom = 24;
    expect(previewChanges([device(24), device(28), device(16)], defaultStandard(), after))
      .toEqual([{ part: 'ram', from: 'Moderate', to: 'Optimal', count: 2 }]);
  });

  it('is empty when the standard did not change', () => {
    expect(previewChanges([device(8)], defaultStandard(), defaultStandard())).toEqual([]);
  });
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `npx vitest run src/features/devices/standards`
Expected: FAIL for the three new files.

- [ ] **Step 3: Implement**

`src/features/devices/standards/gradeColors.js`:

```js
import { GRADES, UNKNOWN } from './defaultStandard.js';

/** The CSS name of each grade: `.g-critical`, `--grade-critical`, … */
export const GRADE_SLUG = {
  Critical: 'critical', 'Needs attention': 'attention', Moderate: 'moderate', Optimal: 'optimal', Unknown: 'unknown',
};

const DARK = '#101828';
const LIGHT = '#ffffff';

function luminance(hex) {
  const n = parseInt(String(hex).slice(1), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255]
    .map((v) => { const c = v / 255; return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4; })
    .reduce((sum, c, i) => sum + c * [0.2126, 0.7152, 0.0722][i], 0);
}

export function contrastRatio(a, b) {
  const [x, y] = [luminance(a), luminance(b)];
  return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05);
}

/** Text on a grade chip: whichever of white or near-black reads better. */
export function inkFor(hex) {
  return contrastRatio(hex, LIGHT) >= contrastRatio(hex, DARK) ? LIGHT : DARK;
}

/** Two of the four grades too alike to tell apart (contrast under 1.5). */
export function tooSimilar(colors) {
  for (let i = 0; i < GRADES.length; i += 1) {
    for (let j = i + 1; j < GRADES.length; j += 1) {
      if (contrastRatio(colors[GRADES[i]], colors[GRADES[j]]) < 1.5) return true;
    }
  }
  return false;
}

export function gradeCssVars(colors) {
  const vars = {};
  for (const grade of [...GRADES, UNKNOWN]) {
    const slug = GRADE_SLUG[grade];
    vars[`--grade-${slug}`] = colors[grade];
    vars[`--grade-${slug}-ink`] = inkFor(colors[grade]);
  }
  return vars;
}
```

`src/features/devices/standards/diffStandard.js`:

```js
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
```

Check against the test's expected order: profiles heavy → desk → mobile, then departments, then colours. In the test, `Engineering RAM optimal from 32 → 24` comes first and `Desk Windows 10 (supported) …` second, which matches.

`src/features/devices/standards/previewChanges.js`:

```js
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
```

- [ ] **Step 4: Run them to see them pass**

Run: `npx vitest run src/features/devices/standards`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/standards
git commit -m "Work out grade text colours, the change summary and which machines a new standard would move

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: The standards list in SharePoint

**Files:**
- Create: `src/features/devices/sharepoint/standardsSchema.js`, `readStandards.js`, `saveStandard.js`, and `standards.test.js` (one test file for all three)

**Interfaces:**
- Consumes: `validateStandard`, `DEFAULT_STANDARD`; `spFetch`, `listPath`, `ITEM_ACCEPT` (`src/features/sharepoint/spClient.js`); `withRetry` (`writePool.js`); `provisionSchema` (`src/features/sharepoint/provision.js`). `provisionSchema(siteUrl, token, { lists })` returns a form digest.
- Produces:
  - `STANDARDS_LIST_NAME`, `STANDARDS_COLUMNS`, `toStandardItem`, `fromStandardItem`
  - `pickStandard(rows)` → `{ standard, version, savedBy, savedOn, history, note }`
  - `readStandards(siteUrl, token)` → the same shape
  - `hasPermission(perms, kind)`, `canEditStandards(siteUrl, token)` → `boolean`
  - `saveStandard({ siteUrl, token, standard, version, summary, savedBy, now })`

- [ ] **Step 1: Write the failing tests**

`src/features/devices/sharepoint/standards.test.js`:

```js
import { describe, it, expect, afterEach, vi } from 'vitest';
import { toStandardItem, fromStandardItem } from './standardsSchema.js';
import { pickStandard, readStandards, hasPermission, canEditStandards } from './readStandards.js';
import { saveStandard } from './saveStandard.js';
import { defaultStandard, DEFAULT_STANDARD } from '../standards/defaultStandard.js';

const reply = (body, status = 200) => ({ ok: status < 300, status, json: async () => body, text: async () => 'x', headers: { get: () => null } });
const row = (version, standard) => ({ Id: version, ...toStandardItem({ version, standard, savedBy: 'aisyah', summary: `v${version}`, savedOn: Date.UTC(2026, 9, version) }) });

describe('standards rows', () => {
  it('round-trip a standard through a list item', () => {
    const back = fromStandardItem(row(4, defaultStandard()));
    expect(back).toMatchObject({ id: 4, version: 4, savedBy: 'aisyah', summary: 'v4', savedOn: Date.UTC(2026, 9, 4) });
    expect(back.standard).toEqual(defaultStandard());
  });

  it('read an unparsable row as no standard', () => {
    expect(fromStandardItem({ Id: 2, Title: 'v2', Standard: '{nope' }).standard).toBeNull();
  });
});

describe('pickStandard', () => {
  it('uses the newest valid version', () => {
    const picked = pickStandard([row(5, defaultStandard()), row(4, defaultStandard())].map(fromStandardItem));
    expect(picked.version).toBe(5);
    expect(picked.note).toBeNull();
    expect(picked.history).toHaveLength(2);
  });

  it('skips a broken newest version and says so', () => {
    const bad = { ...defaultStandard(), schema: 9 };
    const picked = pickStandard([row(6, bad), row(5, defaultStandard())].map(fromStandardItem));
    expect(picked.version).toBe(5);
    expect(picked.note).toBe('Version 6 could not be used, so version 5 is in force.');
  });

  it('falls back to the default standard with nothing saved', () => {
    expect(pickStandard([])).toMatchObject({ standard: DEFAULT_STANDARD, version: 0, note: null });
  });
});

describe('reading and permissions', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('reads the newest twenty, newest first', async () => {
    let asked = '';
    vi.stubGlobal('fetch', async (url) => { asked = decodeURIComponent(url); return reply({ d: { results: [row(2, defaultStandard())] } }); });
    const result = await readStandards('https://x', 't');
    expect(asked).toContain('$orderby=Id desc&$top=20');
    expect(result.version).toBe(2);
  });

  it('treats a list that does not exist yet as the default standard', async () => {
    vi.stubGlobal('fetch', async () => reply({}, 404));
    expect((await readStandards('https://x', 't')).version).toBe(0);
  });

  it('reads a permission bit', () => {
    expect(hasPermission({ Low: '2', High: '0' }, 2)).toBe(true);
    expect(hasPermission({ Low: '1', High: '0' }, 2)).toBe(false);
    expect(hasPermission({ Low: String(0x800), High: '0' }, 12)).toBe(true);
  });

  it('may edit when the list grants AddListItems', async () => {
    vi.stubGlobal('fetch', async () => reply({ d: { EffectiveBasePermissions: { Low: '3', High: '0' } } }));
    expect(await canEditStandards('https://x', 't')).toBe(true);
  });

  it('falls back to "may I create lists" before the list exists', async () => {
    vi.stubGlobal('fetch', async (url) => (url.includes("getByTitle")
      ? reply({}, 404)
      : reply({ d: { EffectiveBasePermissions: { Low: '0', High: '0' } } })));
    expect(await canEditStandards('https://x', 't')).toBe(false);
  });
});

describe('saveStandard', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('refuses an invalid standard before touching SharePoint', async () => {
    const fetch = vi.fn();
    vi.stubGlobal('fetch', fetch);
    const bad = defaultStandard();
    bad.profiles.heavy.ram.criticalBelow = 99;
    await expect(saveStandard({ siteUrl: 'https://x', token: 't', standard: bad, version: 3, summary: 's', savedBy: 'me' }))
      .rejects.toThrow(/in order/);
    expect(fetch).not.toHaveBeenCalled();
  });

  it('turns a 403 into a sentence about edit rights', async () => {
    vi.stubGlobal('fetch', async (url, init = {}) => {
      if (url.endsWith('/_api/contextinfo')) return reply({ d: { GetContextWebInformation: { FormDigestValue: 'D' } } });
      if (url.includes('/fields?')) return reply({ d: { results: [] } });
      if (init.method === 'POST' && url.includes('/items')) return reply({}, 403);
      return reply({ d: {} });
    });
    await expect(saveStandard({ siteUrl: 'https://x', token: 't', standard: defaultStandard(), version: 3, summary: 's', savedBy: 'me' }))
      .rejects.toThrow('You do not have edit rights on the IT Device Standards list. Ask IT.');
  });
});
```

The 403 test drives `provisionSchema` through the fake. If `provisionSchema`'s request sequence makes the fake answer unexpectedly, mock the module at the top of the test file instead, and say so in the report:

```js
vi.mock('../../sharepoint/provision.js', () => ({ provisionSchema: async () => 'D' }));
```

- [ ] **Step 2: Run them to see them fail**

Run: `npx vitest run src/features/devices/sharepoint/standards.test.js`
Expected: FAIL. The modules don't exist.

- [ ] **Step 3: Implement**

`src/features/devices/sharepoint/standardsSchema.js`:

```js
/**
 * Every saved standard is a new row; nothing is ever updated or deleted, so the
 * list IS the history and "restore" is just saving an old one again.
 */
export const STANDARDS_LIST_NAME = 'IT Device Standards';

export const STANDARDS_COLUMNS = [
  { StaticName: 'Standard', Title: 'Standard', kind: 'note' },
  { StaticName: 'SavedOn', Title: 'Saved On', kind: 'datetime' },
  { StaticName: 'SavedBy', Title: 'Saved By', kind: 'text' },
  { StaticName: 'Summary', Title: 'Summary', kind: 'note' },
];

export function toStandardItem({ version, standard, savedBy, summary, savedOn }) {
  return {
    Title: `v${version}`,
    Standard: JSON.stringify(standard),
    SavedOn: new Date(savedOn).toISOString(),
    SavedBy: savedBy ?? '',
    Summary: summary ?? '',
  };
}

export function fromStandardItem(row) {
  let standard = null;
  try {
    standard = JSON.parse(row.Standard);
  } catch {
    standard = null;
  }
  const version = Number(String(row.Title ?? '').replace(/^v/, ''));
  return {
    id: row.Id ?? row.ID ?? null,
    version: Number.isFinite(version) && version > 0 ? version : null,
    standard,
    savedBy: row.SavedBy || null,
    savedOn: row.SavedOn ? new Date(row.SavedOn).getTime() : null,
    summary: row.Summary || null,
  };
}
```

`src/features/devices/sharepoint/readStandards.js`:

```js
import { spFetch, listPath } from '../../sharepoint/spClient.js';
import { STANDARDS_LIST_NAME, fromStandardItem } from './standardsSchema.js';
import { validateStandard } from '../standards/validateStandard.js';
import { DEFAULT_STANDARD } from '../standards/defaultStandard.js';

/**
 * The standard in force: the newest row that is valid. A broken newest row is
 * skipped AND named, never silently -- grading must not stop because somebody
 * saved something odd, and nobody should wonder why their change is not showing.
 */
export function pickStandard(rows) {
  const history = rows.map((row) => ({ ...row, valid: Boolean(row.standard) && validateStandard(row.standard).ok }));
  const inForce = history.find((row) => row.valid) ?? null;
  const skipped = history.length && history[0] !== inForce ? history[0] : null;
  return {
    standard: inForce ? inForce.standard : DEFAULT_STANDARD,
    version: inForce?.version ?? 0,
    savedBy: inForce?.savedBy ?? null,
    savedOn: inForce?.savedOn ?? null,
    history,
    note: skipped
      ? `Version ${skipped.version ?? '?'} could not be used, so ${inForce ? `version ${inForce.version}` : 'the default standard'} is in force.`
      : null,
  };
}

export async function readStandards(siteUrl, token) {
  const response = await spFetch(siteUrl, `${listPath(STANDARDS_LIST_NAME)}/items?$orderby=Id%20desc&$top=20`, { token });
  if (response.status === 404) return pickStandard([]);
  if (!response.ok) throw new Error(`Could not read the device standards (${response.status})`);
  const data = await response.json();
  return pickStandard((data.d?.results ?? []).map(fromStandardItem));
}

/** SP.PermissionKind values: 2 = AddListItems, 12 = ManageLists. */
const ADD_LIST_ITEMS = 2;
const MANAGE_LISTS = 12;

export function hasPermission(perms, kind) {
  const bit = kind - 1;
  const word = bit < 32 ? Number(perms?.Low ?? 0) : Number(perms?.High ?? 0);
  return ((word >>> (bit % 32)) & 1) === 1;
}

const unwrap = (data) => data?.d?.EffectiveBasePermissions ?? data?.d ?? data;

/**
 * May the signed-in user save a standard? SharePoint decides, not the portal:
 * edit rights on the list itself, or -- before the list exists -- the right to
 * create lists, which is what the first save needs.
 */
export async function canEditStandards(siteUrl, token) {
  const onList = await spFetch(siteUrl, `${listPath(STANDARDS_LIST_NAME)}/EffectiveBasePermissions`, { token });
  if (onList.ok) return hasPermission(unwrap(await onList.json()), ADD_LIST_ITEMS);
  if (onList.status !== 404) return false;
  const onSite = await spFetch(siteUrl, '/_api/web/EffectiveBasePermissions', { token });
  if (!onSite.ok) return false;
  return hasPermission(unwrap(await onSite.json()), MANAGE_LISTS);
}
```

`src/features/devices/sharepoint/saveStandard.js`:

```js
import { spFetch, listPath, ITEM_ACCEPT } from '../../sharepoint/spClient.js';
import { withRetry } from '../../sharepoint/writePool.js';
import { provisionSchema } from '../../sharepoint/provision.js';
import { STANDARDS_LIST_NAME, STANDARDS_COLUMNS, toStandardItem } from './standardsSchema.js';
import { validateStandard } from '../standards/validateStandard.js';

/** Append a version. Never updates a row: the list is the history. */
export async function saveStandard({
  siteUrl, token, standard, version, summary, savedBy, now = Date.now(),
}) {
  const check = validateStandard(standard);
  if (!check.ok) throw new Error(check.errors[0].message);

  const digest = await provisionSchema(siteUrl, token, {
    lists: [{
      title: STANDARDS_LIST_NAME,
      description: 'Versions of what counts as Critical, Needs attention, Moderate and Optimal for each part',
      columns: STANDARDS_COLUMNS,
    }],
  });

  const response = await withRetry(() => spFetch(siteUrl, `${listPath(STANDARDS_LIST_NAME)}/items`, {
    token, digest, method: 'POST', accept: ITEM_ACCEPT,
    body: toStandardItem({ version, standard, savedBy, summary, savedOn: now }),
  }));
  if (response.status === 403) throw new Error('You do not have edit rights on the IT Device Standards list. Ask IT.');
  if (!response.ok) throw new Error(`Could not save the standard (${response.status})`);
}
```

Before relying on it, check `provisionSchema`'s list declaration keys in `src/features/sharepoint/provision.js` (`title`, `description`, `columns`; `provisionLists.js` uses the same shape). Also check that it returns the digest.

- [ ] **Step 4: Run them to see them pass**

Run: `npx vitest run src/features/devices/sharepoint`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/sharepoint
git commit -m "Keep every saved device standard as a version in SharePoint, and let SharePoint say who may save one

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: Load the standard, grade with it, and colour by it

**Files:**
- Create: `src/features/devices/useStandards.js`
- Modify: `src/features/devices/useDevices.js`, `src/pages/DevicesPage.jsx`, `src/pages/DeviceDetailPage.jsx`, `src/styles/devices.css`

**Interfaces:**
- Consumes: `readStandards`, `canEditStandards`, `pickStandard` (Task 4); `regrade` (Task 2); `gradeCssVars` (Task 3).
- Produces:
  - `useStandards()` → `{ standard, version, savedBy, savedOn, history, note, canEdit, loading, error, reload }`
  - `useDevices()` → `{ devices (graded with the loaded standard), loading, error, reload, standards }`
  - The CSS classes `.g-critical`, `.g-attention`, `.g-moderate`, `.g-optimal` and `.g-unknown`, and the root class `.dv-graded`

- [ ] **Step 1: Write `useStandards`**

Create `src/features/devices/useStandards.js`:

```js
import { useCallback, useEffect, useState } from 'react';
import { useIsAuthenticated } from '@azure/msal-react';
import { useSharePointToken } from '../../hooks/useRequests';
import { readStandards, canEditStandards, pickStandard } from './sharepoint/readStandards';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

/**
 * The standard in force, and whether this person may change it. Until it has
 * loaded -- or if it cannot be read -- the default standard grades the fleet,
 * so a page is never ungraded and never blocked on this read.
 */
export function useStandards() {
  const isAuthenticated = useIsAuthenticated();
  const getToken = useSharePointToken();
  const [state, setState] = useState({ loaded: pickStandard([]), canEdit: false, loading: true, error: '' });
  const [nonce, setNonce] = useState(0);
  const reload = useCallback(() => setNonce((n) => n + 1), []);

  useEffect(() => {
    if (!isAuthenticated) return undefined;
    let cancelled = false;
    (async () => {
      try {
        const tokenRes = await getToken();
        const [loaded, canEdit] = await Promise.all([
          readStandards(SHAREPOINT_SITE_URL, tokenRes.accessToken),
          canEditStandards(SHAREPOINT_SITE_URL, tokenRes.accessToken).catch(() => false),
        ]);
        if (!cancelled) setState({ loaded, canEdit, loading: false, error: '' });
      } catch {
        if (!cancelled) {
          setState((current) => ({
            ...current, loading: false, error: 'Using the default standard: the saved one could not be read.',
          }));
        }
      }
    })();
    return () => { cancelled = true; };
  }, [isAuthenticated, getToken, nonce]);

  return { ...state.loaded, canEdit: state.canEdit, loading: state.loading, error: state.error, reload };
}
```

The `setState` calls sit inside the async callback, the same pattern as `useDevices.js`, so `react-hooks/set-state-in-effect` does not fire. If lint disagrees, mirror `useDevices.js` exactly.

- [ ] **Step 2: Grade in `useDevices`**

In `src/features/devices/useDevices.js`:
- Change the React import to include `useMemo`.
- Add these imports:

```js
import { useStandards } from './useStandards';
import { regrade } from './derive/regrade';
```

- Replace the final `return` with:

```js
  const standards = useStandards();
  // Graded on the way out with the standard in force: saving a standard
  // regrades every machine on the next render, with no re-scan.
  const graded = useMemo(() => regrade(devices, standards.standard), [devices, standards.standard]);

  return { devices: graded, loading, error, reload, standards };
```

`useStandards()` must be called unconditionally, alongside the other hooks at the top of `useDevices` and before any early return. Move the call up if needed.

- [ ] **Step 3: Set the grade colours on the device pages**

In both `src/pages/DevicesPage.jsx` and `src/pages/DeviceDetailPage.jsx`:
- Read `standards` from `useDevices()`, e.g. `const { devices: saved, loading, error, reload, standards } = useDevices();` on DevicesPage, and the same destructure on DeviceDetailPage.
- Import `gradeCssVars` from `../features/devices/standards/gradeColors`.
- Wrap EVERYTHING inside `<AppShell …>` in:

```jsx
<div className="dv-graded" style={gradeCssVars(standards.standard.colors)}>
  {/* existing children */}
</div>
```

- Directly inside that wrapper, before the existing content, add:

```jsx
{(standards.note || standards.error) && (
  <p className="dv-standard-note" role="status">{standards.note || standards.error}</p>
)}
```

Append to `src/styles/devices.css`:

```css
/* --- Grade colours: set from the saved standard on .dv-graded ----------- */
.dv-graded {
  --grade-critical: #dc2626; --grade-critical-ink: #ffffff;
  --grade-attention: #f59e0b; --grade-attention-ink: #101828;
  --grade-moderate: #1a88de; --grade-moderate-ink: #ffffff;
  --grade-optimal: #12a150; --grade-optimal-ink: #ffffff;
  --grade-unknown: #8a97a8; --grade-unknown-ink: #101828;
}
.g-critical { background: var(--grade-critical); color: var(--grade-critical-ink); }
.g-attention { background: var(--grade-attention); color: var(--grade-attention-ink); }
.g-moderate { background: var(--grade-moderate); color: var(--grade-moderate-ink); }
.g-optimal { background: var(--grade-optimal); color: var(--grade-optimal-ink); }
.g-unknown { background: var(--grade-unknown); color: var(--grade-unknown-ink); }
.dv-standard-note { margin: 0 0 12px; padding: 8px 12px; border-radius: 10px; background: var(--it-accent-wash); color: var(--it-ink); font-size: 14px; }
```

The inline `style` on `.dv-graded` overrides these fallbacks with the saved colours.

- [ ] **Step 4: Check**

Run: `npx eslint src`, `npm run build`, `npm test`
Expected: only the two baseline lint errors, a passing build, and green tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/useStandards.js src/features/devices/useDevices.js src/pages/DevicesPage.jsx src/pages/DeviceDetailPage.jsx src/styles/devices.css
git commit -m "Grade the fleet with the saved standard on every read, and colour the device section with its grade colours

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: The map shows a bar per part

**Files:**
- Create: `src/features/devices/ui/PartBars.jsx`
- Modify:
  - `src/features/devices/map/zones.js`, `src/features/devices/map/zones.test.js`
  - `src/features/devices/ui/MapTile.jsx`, `DepartmentView.jsx`, `LocationView.jsx`, `DeviceMap.jsx`
  - `src/pages/DevicesPage.jsx`, `src/styles/devices.css`
  - `src/features/devices/derive/persona.js` and `persona.test.js` (drop `personaFor` / `BY_DEPARTMENT`)

**Interfaces:**
- Consumes: `PARTS`, `GRADES`, `UNKNOWN`, `profileKeyFor`, `PROFILE_SHORT` (Task 1); `GRADE_SLUG` (Task 3); the graded device fields `parts`, `criticalParts`, `attentionParts` (Task 2).
- Produces:
  - `summarise(name, devices)` → `{ name, count, laptops, desktops, other, parts, criticalMachines, attentionMachines, partBars: [{ key, label, segs: [{ grade, share }] }] }`
  - The fields `fit`, `critical`, `attention` and `health` are REMOVED.
  - `machinesIn` sorts by number of critical parts, then attention parts, then name.
  - `<PartBars bars />`
  - `<DeviceMap … standard />`, which passes `standard` down to `<LocationView standard />`.

- [ ] **Step 1: Rewrite the zones tests that read the old fields**

In `src/features/devices/map/zones.test.js`:
- Change the `m()` helper so it builds graded devices: drop `fitStatus: 'Optimal'`, and add `parts` and the part arrays through `partGrades`:

```js
import { partGrades } from '../derive/partGrades.js';
import { defaultStandard } from '../standards/defaultStandard.js';

const S = defaultStandard();
const m = (over) => {
  const base = {
    id: Math.random(), computerName: 'PC', owner: 'A', location: 'F1', department: 'ENGINEERING',
    deviceType: 'Laptop', status: null, scanComplete: true,
    cpuGenerationRank: 12, installedRamGB: 32, storageTotalGB: 512, storageType: 'SSD only',
    dedicatedGpu: true, osSupported: true, windowsMajor: 11, ...over,
  };
  return { ...base, ...partGrades(base, S) };
};
```

- Replace the two `summarise` tests with:

```js
describe('summarise', () => {
  it('counts laptops, desktops and anything else apart', () => {
    const s = summarise('F1', [m(), m({ deviceType: 'Desktop' }), m({ deviceType: 'Unknown' })]);
    expect(s).toMatchObject({ count: 3, laptops: 1, desktops: 1, other: 1 });
  });

  it('counts each part\u2019s grades on its own, and the machines with a critical part', () => {
    const s = summarise('F1', [m({ installedRamGB: 4 }), m({ storageType: 'HDD only', installedRamGB: 4 }), m()]);
    expect(s.parts.ram.Critical).toBe(2);
    expect(s.parts.storage.Critical).toBe(1);
    expect(s.criticalMachines).toBe(2);
    expect(s.partBars.find((bar) => bar.key === 'ram').segs).toEqual([
      { grade: 'Critical', share: 2 / 3 }, { grade: 'Optimal', share: 1 / 3 },
    ]);
  });
});
```

- In the `machinesIn` "worst first" test, replace `fitStatus: 'Critical'` with `installedRamGB: 4` so machine `A` has a critical part.

- [ ] **Step 2: Run them to see them fail**

Run: `npx vitest run src/features/devices/map/zones.test.js`
Expected: FAIL. `parts` and `criticalMachines` are undefined.

- [ ] **Step 3: Implement**

In `src/features/devices/map/zones.js`:
- Add the import `import { GRADES, UNKNOWN, PARTS } from '../standards/defaultStandard.js';`
- Replace `LEVELS` and `summarise` with:

```js
const ALL_GRADES = [...GRADES, UNKNOWN];

export function summarise(name, devices) {
  const parts = Object.fromEntries(PARTS.map(({ key }) => [key, Object.fromEntries(ALL_GRADES.map((g) => [g, 0]))]));
  let laptops = 0;
  let desktops = 0;
  let other = 0;
  let criticalMachines = 0;
  let attentionMachines = 0;

  for (const device of devices) {
    for (const { key } of PARTS) {
      const grade = device.parts?.[key]?.grade;
      parts[key][ALL_GRADES.includes(grade) ? grade : UNKNOWN] += 1;
    }
    if (device.criticalParts?.length) criticalMachines += 1;
    else if (device.attentionParts?.length) attentionMachines += 1;
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
    parts,
    criticalMachines,
    attentionMachines,
    partBars: PARTS.map(({ key, label }) => ({
      key,
      label,
      segs: count
        ? ALL_GRADES.map((grade) => ({ grade, share: parts[key][grade] / count })).filter((seg) => seg.share > 0)
        : [],
    })),
  };
}
```

- Replace `FIT_RANK` and the sort in `machinesIn` with:

```js
  return [...rows].sort((a, b) =>
    (b.criticalParts?.length ?? 0) - (a.criticalParts?.length ?? 0)
    || (b.attentionParts?.length ?? 0) - (a.attentionParts?.length ?? 0)
    || String(a.computerName ?? '').localeCompare(String(b.computerName ?? '')));
```

Create `src/features/devices/ui/PartBars.jsx`:

```jsx
import { GRADE_SLUG } from '../standards/gradeColors';

/** Five thin bars, one per part, each split by how many machines hold each grade on it. */
export default function PartBars({ bars }) {
  return (
    <span className="pb" aria-hidden="true">
      {bars.map((bar) => (
        <span key={bar.key} className="pb-row">
          <span className="pb-label">{bar.label}</span>
          <span className="pb-track">
            {bar.segs.map((seg) => (
              <span key={seg.grade} className={`pb-seg g-${GRADE_SLUG[seg.grade]}`} style={{ width: `${seg.share * 100}%` }} />
            ))}
          </span>
        </span>
      ))}
    </span>
  );
}
```

In `src/features/devices/ui/MapTile.jsx`:
- Remove `LEVEL_CLASS`, and import `PartBars` from `./PartBars` and `PARTS` from `../standards/defaultStandard`.
- Replace the `label` const with:

```jsx
  const criticalOn = PARTS.filter(({ key }) => summary.parts?.[key]?.Critical)
    .map(({ key, label: part }) => `${part} critical on ${summary.parts[key].Critical}`);
  const label = [
    title, `${summary.count} machines`, `${summary.laptops} laptops`, `${summary.desktops} desktops`, ...criticalOn,
  ].join(', ');
```

- Replace the badge expression with:

```jsx
        {variant === 'zone' && (summary.criticalMachines > 0
          ? <span className="mt-badge g-critical">{summary.criticalMachines} with a critical part</span>
          : summary.count > 0 && summary.attentionMachines === 0 && <span className="mt-badge g-optimal">All clear</span>)}
```

- Replace the `{summary.health.length > 0 && (…)}` block with `{summary.count > 0 && <PartBars bars={summary.partBars} />}`.

In `src/features/devices/ui/DepartmentView.jsx`, replace the two header stats that read `summary.critical` and `summary.attention` with:

```jsx
          {!place && <span><strong>{summary.criticalMachines}</strong> with a critical part</span>}
          {!place && <span><strong>{summary.attentionMachines}</strong> needing attention</span>}
```

Then add `<PartBars bars={summary.partBars} />` after `.dv-level-stats` (import it).

In `src/features/devices/ui/DeviceMap.jsx`:
- Accept a `standard` prop and pass it to `<LocationView standard={standard} … />`.

In `src/features/devices/ui/LocationView.jsx`:
- Accept `standard`.
- Replace the `personaFor` import with `import { profileKeyFor } from '../standards/defaultStandard';`
- Change the eyebrow to:

```jsx
          eyebrow={tile.name === 'Unassigned' ? 'Department' : standard.profiles[profileKeyFor(standard, tile.name)].label}
```

In `src/pages/DevicesPage.jsx`, render `<DeviceMap devices={saved} loading={loading} params={params} standard={standards.standard} />`.

In `src/features/devices/derive/persona.js`, remove `BY_DEPARTMENT` and `personaFor`, since the mapping now lives in the standard. Keep `PERSONAS`. Remove the tests for `personaFor` from `persona.test.js`. If that leaves the file empty, add one test that `PERSONAS.HEAVY.label` is `'Engineering / Technical / Media'`. Run `grep -rn "personaFor" src` and fix any other importer.

Append to `src/styles/devices.css`. Also remove the now-unused `.mt-health` and `.mt-seg*` rules, and `.mt-badge-crit` / `.mt-badge-ok`.

```css
/* --- Part bars on a map tile ------------------------------------------- */
.pb { display: flex; flex-direction: column; gap: 4px; }
.pb-row { display: grid; grid-template-columns: 64px 1fr; align-items: center; gap: 8px; font-size: 12px; }
.pb-label { color: var(--it-ink-soft); font-weight: 600; }
.mt-hub .pb-label { color: var(--it-on-brand-dim); }
.pb-track { display: flex; height: 6px; border-radius: 999px; overflow: hidden; background: var(--it-line); }
.pb-seg { display: block; height: 100%; }
.dv-level .pb-label { color: var(--it-on-brand-dim); }
```

- [ ] **Step 4: Check**

Run: `npx vitest run src/features/devices`, `npx eslint src`, `npm run build`
Expected: PASS, only the baseline lint errors, and a passing build.

- [ ] **Step 5: Commit**

```bash
git add -A src/features/devices src/pages/DevicesPage.jsx src/styles/devices.css
git commit -m "Show a bar per part on every map tile, and list a department's machines by how many parts are critical

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 7: Part chips on the cards, and a part table on the machine page

**Files:**
- Create: `src/features/devices/ui/PartChips.jsx`, `src/features/devices/ui/PartTable.jsx`
- Modify: `src/features/devices/ui/MachineCard.jsx`, `src/pages/DeviceDetailPage.jsx`, `src/styles/devices.css`

**Interfaces:**
- Consumes: `PARTS`, `GRADE_SLUG`; the graded fields.
- Produces: `<PartChips device />` and `<PartTable device />`.

- [ ] **Step 1: Create the components**

`src/features/devices/ui/PartChips.jsx`:

```jsx
import { PARTS } from '../standards/defaultStandard';
import { GRADE_SLUG } from '../standards/gradeColors';

/** One chip per part, coloured by that part's own grade. */
export default function PartChips({ device }) {
  return (
    <span className="pc">
      {PARTS.map(({ key, label }) => {
        const part = device.parts?.[key];
        const grade = part?.grade ?? 'Unknown';
        return (
          <span key={key} className={`pc-chip g-${GRADE_SLUG[grade]}`} title={part?.reason}>
            {label} · {grade}
          </span>
        );
      })}
    </span>
  );
}
```

`src/features/devices/ui/PartTable.jsx`:

```jsx
import { Card } from '../../../components/ui/Surfaces';
import { PARTS } from '../standards/defaultStandard';
import { GRADE_SLUG } from '../standards/gradeColors';

/** Every part against the standard for this machine's desk, with the reason for each grade. */
export default function PartTable({ device }) {
  return (
    <Card className="dd-fit">
      <h2 className="dd-group-title">
        Parts against the {device.personaLabel} standard
        <span className="dd-group-hint">{device.personaBlurb}</span>
      </h2>
      <div className="pt-scroll">
        <table className="pt">
          <thead>
            <tr><th scope="col">Part</th><th scope="col">This machine</th><th scope="col">Grade</th><th scope="col">Why</th></tr>
          </thead>
          <tbody>
            {PARTS.map(({ key, label }) => {
              const part = device.parts?.[key] ?? { grade: 'Unknown', value: '—', reason: '' };
              return (
                <tr key={key}>
                  <th scope="row">{label}</th>
                  <td>{part.value}</td>
                  <td><span className={`pc-chip g-${GRADE_SLUG[part.grade]}`}>{part.grade}</span></td>
                  <td>{part.reason}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <dl className="dd-fit-facts">
        <div><dt>Action</dt><dd>{device.actionRequired ?? '—'}</dd></div>
        <div>
          <dt>Suggested form factor</dt>
          <dd>{device.suggestedFormFactor ?? '—'}<span className="dd-fit-note">{device.formFactorNote}</span></dd>
        </div>
        <div>
          <dt>Office licence</dt>
          <dd>{device.licenseStatus ?? '—'}<span className="dd-fit-note">{device.licenseNote}</span></dd>
        </div>
        <div>
          <dt>Server link</dt>
          <dd>
            {device.serverDependent ? device.networkRisk : 'Not server-bound'}
            <span className="dd-fit-note">{device.networkNote}</span>
          </dd>
        </div>
      </dl>
    </Card>
  );
}
```

- [ ] **Step 2: Use them**

In `src/features/devices/ui/MachineCard.jsx`:
- Remove `TONE` and `tone`.
- Import `PartChips` from `./PartChips` and `GRADE_SLUG` from `../standards/gradeColors`.
- The link's className becomes `"mc"`, and its `aria-label` becomes:

```jsx
`${who}, ${device.computerName}, ${(device.criticalParts ?? []).length} critical part${(device.criticalParts ?? []).length === 1 ? '' : 's'}`
```

- Give each bar its part's grade. Change `bars()` entries to carry `part: 'cpu' | 'ram' | 'storage'`, and the fill to:

```jsx
<span className={`mc-bar-fill g-${GRADE_SLUG[device.parts?.[bar.part]?.grade ?? 'Unknown']}`} style={{ width: `${Math.min(1, bar.share) * 100}%` }} />
```

- Replace the `mc-verdict` span with `<PartChips device={device} />`.

In `src/pages/DeviceDetailPage.jsx`:
- Remove `FIT_TONE`.
- Import `PartTable` from `../features/devices/ui/PartTable`.
- Replace the whole `<Card className="dd-fit"> … </Card>` block with `<PartTable device={device} />`.

Append to `src/styles/devices.css`. Also remove the now-unused `.mc-crit/.mc-attn/.mc-mod/.mc-ok` border rules, the `.mc-… .mc-bar-fill` colour rules, and `.mc-verdict*`.

```css
/* --- Part chips and the part table ------------------------------------- */
.pc { display: flex; flex-wrap: wrap; gap: 4px; }
.pc-chip { display: inline-flex; align-items: center; min-height: 24px; padding: 2px 8px; border-radius: 999px; font-size: 12px; font-weight: 700; white-space: nowrap; }
.mc .mc-bar-fill { background: var(--grade-unknown); }
.mc .mc-bar-fill[class*="g-"] { color: transparent; }
.pt-scroll { overflow-x: auto; }
.pt { width: 100%; border-collapse: collapse; font-size: 14px; }
.pt th, .pt td { text-align: left; padding: 10px 12px; border-bottom: 1px solid var(--it-line); vertical-align: top; }
.pt thead th { font-size: 12px; letter-spacing: .08em; text-transform: uppercase; color: var(--it-ink-soft); }
```

The `.mc .mc-bar-fill` rule gives a grey fallback when no grade class applies. The `.g-*` classes set the real fill.

- [ ] **Step 3: Check**

Run: `npx eslint src`, `npm run build`, `npm test`
Expected: only the baseline lint errors, a passing build, and green tests.

- [ ] **Step 4: Commit**

```bash
git add src/features/devices/ui src/pages/DeviceDetailPage.jsx src/styles/devices.css
git commit -m "Grade chips on every machine card, and a part-by-part table on the machine page

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 8: Dashboard, register and filters by part

**Files:**
- Modify:
  - `src/features/devices/stats/deviceStats.js` and its test
  - `src/features/devices/deviceFilters.js` and its test
  - `src/features/devices/ui/DeviceTable.jsx`, `DeviceCharts.jsx`, `DepartmentHeatmap.jsx`
  - `src/pages/DevicesPage.jsx`

**Interfaces:**
- Consumes: `PARTS`, `GRADE_SLUG`, and the graded fields.
- Produces:
  - `complianceSummary(devices)` keeps `critical` and `criticalPct`, which now mean machines with a critical part out of machines with at least one graded part. It drops `optimal` and `graded`, and gains `judged`.
  - `criticalPartsByDepartment(devices)` → `[{ department, persona, total, criticalMachines, cpu, ram, storage, graphics, windows }]`, sorted worst first.
  - Filters `part=<key>:<grade>` and `critical=1`. The `fit` filter is removed.
  - `FIT_ORDER` and `fitByDepartment` are REMOVED.

- [ ] **Step 1: Write the failing tests**

In `src/features/devices/stats/deviceStats.test.js`, replace the tests of `complianceSummary`'s `critical` / `criticalPct` / `optimal` and every `fitByDepartment` test with the following. Keep the other `complianceSummary` assertions (licences, server dependency, form factor).

```js
import { criticalPartsByDepartment } from './deviceStats.js';
import { partGrades } from '../derive/partGrades.js';
import { defaultStandard } from '../standards/defaultStandard.js';

const S = defaultStandard();
const graded = (over) => {
  const base = { department: 'HR', scanComplete: true, cpuGenerationRank: 12, installedRamGB: 16, storageTotalGB: 512,
    storageType: 'SSD only', dedicatedGpu: false, osSupported: true, windowsMajor: 11, ...over };
  return { ...base, ...partGrades(base, S) };
};

describe('complianceSummary — critical parts', () => {
  it('counts machines with any critical part, as a share of machines graded at all', () => {
    const s = complianceSummary([graded({ installedRamGB: 4 }), graded(), graded({ scanComplete: false })]);
    expect(s.critical).toBe(1);
    expect(s.judged).toBe(2);
    expect(s.criticalPct).toBe(50);
  });
});

describe('criticalPartsByDepartment', () => {
  it('counts critical machines per part per department, worst department first', () => {
    const rows = criticalPartsByDepartment([
      graded({ department: 'HR', installedRamGB: 4 }), graded({ department: 'HR' }),
      graded({ department: 'SALES', installedRamGB: 4, storageType: 'HDD only' }),
    ]);
    expect(rows[0]).toMatchObject({ department: 'SALES', total: 1, criticalMachines: 1, ram: 1, storage: 1, cpu: 0 });
    expect(rows[1]).toMatchObject({ department: 'HR', total: 2, criticalMachines: 1, ram: 1 });
  });
});
```

`complete()` in `deviceStats.js` may filter out `scanComplete === false` rows. Read it, and if it does, adjust the expected `judged` to match what `complete` keeps; the intent is "machines with at least one graded part".

In `src/features/devices/deviceFilters.test.js`, append:

```js
describe('part filters', () => {
  const rows = [
    { computerName: 'A', parts: { ram: { grade: 'Critical' } }, criticalParts: ['ram'] },
    { computerName: 'B', parts: { ram: { grade: 'Optimal' } }, criticalParts: [] },
  ];
  it('filters by one part\u2019s grade', () => {
    expect(applyFilters(rows, { part: 'ram:Critical' }).map((r) => r.computerName)).toEqual(['A']);
  });
  it('filters to machines with any critical part', () => {
    expect(applyFilters(rows, { critical: '1' }).map((r) => r.computerName)).toEqual(['A']);
  });
  it('ignores an old fit link instead of failing', () => {
    expect(applyFilters(rows, { fit: 'Critical' })).toHaveLength(2);
  });
});
```

Delete any existing `fit:` matcher test.

- [ ] **Step 2: Run them to see them fail**

Run: `npx vitest run src/features/devices/stats src/features/devices/deviceFilters.test.js`
Expected: FAIL.

- [ ] **Step 3: Implement**

In `src/features/devices/stats/deviceStats.js`:
- Import `PARTS` and `UNKNOWN` from `../standards/defaultStandard.js`.
- Delete `FIT_ORDER` and `fitByDepartment`.
- In `complianceSummary`, replace the `graded` const and the `graded`, `critical`, `criticalPct` and `optimal` fields with:

```js
  const judged = rows.filter((device) => PARTS.some(({ key }) => device.parts?.[key] && device.parts[key].grade !== UNKNOWN));
  const withCritical = judged.filter((device) => (device.criticalParts?.length ?? 0) > 0).length;
```

and in the returned object:

```js
    judged: judged.length,
    critical: withCritical,
    criticalPct: pct(withCritical, judged.length),
```

- Add:

```js
/**
 * One row per department: how many of its machines are Critical on each part.
 * Worst first by the share of its machines with any critical part.
 */
export function criticalPartsByDepartment(devices) {
  const byDepartment = new Map();
  for (const device of complete(devices)) {
    const department = labelOf(device.department);
    const row = byDepartment.get(department) ?? {
      department, persona: device.personaLabel ?? 'Unclassified', total: 0, criticalMachines: 0,
      ...Object.fromEntries(PARTS.map(({ key }) => [key, 0])),
    };
    row.total += 1;
    if (device.criticalParts?.length) row.criticalMachines += 1;
    for (const key of device.criticalParts ?? []) row[key] += 1;
    byDepartment.set(department, row);
  }
  return [...byDepartment.values()].sort((a, b) =>
    (b.criticalMachines / b.total) - (a.criticalMachines / a.total)
    || b.total - a.total || a.department.localeCompare(b.department));
}
```

In `src/features/devices/deviceFilters.js`, replace the `fit:` matcher with:

```js
  part: (device, value) => {
    const [key, grade] = String(value).split(':');
    return device.parts?.[key]?.grade === grade;
  },
  critical: (device) => (device.criticalParts?.length ?? 0) > 0,
```

In `src/features/devices/ui/DepartmentHeatmap.jsx`, replace the component with a department × part grid. Keep the existing `.dv-heat*` classes where they fit:

```jsx
import { Card, EmptyState } from '../../../components/ui/Surfaces';
import { criticalPartsByDepartment } from '../stats/deviceStats';
import { PARTS } from '../standards/defaultStandard';

/**
 * Which department is in the most trouble, and on which part. Each cell is the
 * number of the department's machines Critical on that part; it opens them.
 */
export default function DepartmentHeatmap({ devices, onSelect }) {
  const rows = criticalPartsByDepartment(devices);
  return (
    <Card className="chart-card dv-heat">
      <div className="chart-head">
        <h3>Critical parts by department</h3>
        <p>How many machines in each department are Critical on each part. Click a number to see them.</p>
      </div>
      {rows.length === 0 ? (
        <EmptyState>No devices imported yet.</EmptyState>
      ) : (
        <div className="hm-scroll">
          <table className="hm">
            <thead>
              <tr>
                <th scope="col">Department</th>
                {PARTS.map(({ key, label }) => <th scope="col" key={key}>{label}</th>)}
                <th scope="col">Machines</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => (
                <tr key={row.department}>
                  <th scope="row" title={row.persona}>{row.department}</th>
                  {PARTS.map(({ key, label }) => (
                    <td key={key}>
                      {row[key] > 0 ? (
                        <button
                          type="button"
                          className="hm-cell g-critical"
                          onClick={() => onSelect?.(row.department, key)}
                          aria-label={`${row.department}: ${row[key]} machines with a critical ${label}. Show them.`}
                        >
                          {row[key]}
                        </button>
                      ) : <span className="hm-zero">0</span>}
                    </td>
                  ))}
                  <td>{row.total}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}
```

In `src/features/devices/ui/DeviceCharts.jsx`:
- Remove `FIT_COLOUR`.
- Replace the "Fit for the work" `BarChart` with a "Critical parts" chart: one bar per part, the value being how many devices are Critical on it, with colour `var(--grade-critical)`. Selecting a bar calls `onFilter('part', `${key}:Critical`)`.
- Build the rows in exactly the shape `paint()` returns (read `paint` and `select` at the top of the file), e.g.:

```jsx
  const criticalByPart = PARTS.map(({ key, label }) => ({
    key, label, value: devices.filter((d) => d.parts?.[key]?.grade === 'Critical').length,
  }));
```

  Then map to `paint`'s shape with colour `'var(--grade-critical)'`, and on select find the part by label.

In `src/features/devices/ui/DeviceTable.jsx`:
- In `CALCULATED_COLUMNS`, replace `{ key: 'fitStatus', label: 'Device Health' }` and `{ key: 'fitReasons', label: 'Why' }` with:

```js
  { key: 'gradeCpu', label: 'CPU Grade' },
  { key: 'gradeRam', label: 'RAM Grade' },
  { key: 'gradeStorage', label: 'Storage Grade' },
  { key: 'gradeGraphics', label: 'Graphics Grade' },
  { key: 'gradeWindows', label: 'Windows Grade' },
```

- In `LEAD_KEYS`, replace `'fitStatus'` with `'gradeCpu', 'gradeRam', 'gradeStorage', 'gradeGraphics', 'gradeWindows'`.
- Remove `FIT_TONE`.
- Change the cell-class function so a grade column renders as a chip. Import `GRADE_SLUG` from `../standards/gradeColors`, then:

```js
const GRADE_KEYS = new Set(['gradeCpu', 'gradeRam', 'gradeStorage', 'gradeGraphics', 'gradeWindows']);
// inside the existing class function:
  if (GRADE_KEYS.has(key)) return `dt-grade g-${GRADE_SLUG[device[key]] ?? 'unknown'}`;
```

- In `FILTER_LABELS`, replace `fit: 'Device health'` with `part: 'Part grade', critical: 'Has a critical part'`, and add `'critical'` to `FLAG_FILTERS`.

In `src/pages/DevicesPage.jsx`:
- In `FILTER_KEYS`, replace `'fit'` with `'part', 'critical'`.
- Change the "Not fit for the work" `StatCard`: label `"Machines with a critical part"`, value `compliance.criticalPct ?? '—'`, unit `` `% · ${compliance.critical} machines` ``, `onClick={() => openRegister('critical', '1')}`.
- Change the heatmap `onSelect` to:

```jsx
                onSelect={(name, part) => {
                  setParam('department', name);
                  openRegister('part', `${part}:Critical`);
                }}
```

Append to `src/styles/devices.css`:

```css
/* --- Part heatmap and register grade cells ------------------------------ */
.hm-scroll { overflow-x: auto; }
.hm { width: 100%; border-collapse: collapse; font-size: 14px; }
.hm th, .hm td { padding: 8px 10px; border-bottom: 1px solid var(--it-line); text-align: center; }
.hm th[scope="row"], .hm thead th:first-child { text-align: left; }
.hm-cell { min-width: 44px; min-height: 36px; border: 0; border-radius: 8px; font: inherit; font-weight: 700; cursor: pointer; }
.hm-zero { color: var(--it-ink-soft); }
.dt-grade { font-weight: 700; }
```

- [ ] **Step 4: Check**

Run: `npm test`, `npx eslint src`, `npm run build`. Also run `grep -rn "fitStatus\|fitReasons\|FIT_ORDER\|fitByDepartment\|'Needs Attention'" src`.
Expected: green tests, only the baseline lint errors, a passing build, and NO grep hits.

- [ ] **Step 5: Commit**

```bash
git add -A src/features/devices src/pages/DevicesPage.jsx src/styles/devices.css
git commit -m "Count critical parts on the dashboard, filter the register by a part's grade, and show a department × part heatmap

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 9: The Standards tab

**Files:**
- Create: `src/features/devices/ui/StandardsPage.jsx`
- Modify: `src/pages/DevicesPage.jsx`, `src/styles/devices.css`

**Interfaces:**
- Consumes:
  - `useDevices().standards` (Task 5)
  - Task 1: `cloneStandard`, `PROFILE_KEYS`, `PROFILE_SHORT`, `GRADES`, `UNKNOWN`, `STORAGE_TYPES`, `DEFAULT_COLORS`
  - `validateStandard`
  - Task 3: `diffStandard`, `summaryOf`, `previewChanges`, `inkFor`, `tooSimilar`, `GRADE_SLUG`
  - Task 4: `saveStandard`
  - `useConfirm`, `Button`, `ErrorBanner`, `labelOf`, `UNASSIGNED`, `useSharePointToken`, `formatMYT`
- Produces: `<StandardsPage devices standards />` and the `standards` view of `/devices`.

- [ ] **Step 1: Create the page**

`src/features/devices/ui/StandardsPage.jsx`:

```jsx
import { useMemo, useState } from 'react';
import Button from '../../../components/ui/Button';
import { ErrorBanner } from '../../../components/ui/Surfaces';
import { useConfirm } from '../../../components/ui/useConfirm';
import { useSharePointToken } from '../../../hooks/useRequests';
import { formatMYT } from '../../../utils/malaysiaTime';
import { labelOf, UNASSIGNED } from '../deviceFilters';
import {
  cloneStandard, PROFILE_KEYS, PROFILE_SHORT, GRADES, UNKNOWN, STORAGE_TYPES, DEFAULT_COLORS,
} from '../standards/defaultStandard';
import { validateStandard } from '../standards/validateStandard';
import { diffStandard, summaryOf } from '../standards/diffStandard';
import { previewChanges } from '../standards/previewChanges';
import { inkFor, tooSimilar, GRADE_SLUG } from '../standards/gradeColors';
import { saveStandard } from '../sharepoint/saveStandard';

const SITE = import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const NUMERIC = [
  { part: 'cpu', name: 'CPU generation', unit: 'gen', max: 16, note: 'Intel-equivalent, so AMD lands on the same scale. Pentium, Celeron and AMD before Ryzen are always Critical.' },
  { part: 'ram', name: 'RAM', unit: 'GB', max: 64, note: 'Installed memory (the slots added up), not what Windows reports as usable.' },
  { part: 'storageSize', name: 'Storage size', unit: 'GB', max: 1024, note: 'Total disk, leaving out the scan\u2019s own USB disk. Storage takes the worse of size and type.' },
];
const CUTS = [
  { cut: 'criticalBelow', label: 'Critical below', grade: 'Critical' },
  { cut: 'attentionBelow', label: 'Needs attention below', grade: 'Needs attention' },
  { cut: 'optimalFrom', label: 'Optimal from', grade: 'Optimal' },
];
const CHOICES = [
  { part: 'storageType', name: 'Storage type', note: 'A machine booting from a hard disk waits on it for everything.', rows: STORAGE_TYPES.map((k) => [k, k === 'Mixed' ? 'Mixed (SSD + hard disk)' : k]) },
  { part: 'graphics', name: 'Graphics', note: 'Drawing and rendering need a dedicated card; desk work does not.', rows: [['dedicated', 'Dedicated card'], ['builtIn', 'Built-in graphics']] },
  { part: 'windows', name: 'Windows', note: 'Out of support means no security updates.', rows: [['win11', 'Windows 11'], ['win10Supported', 'Windows 10 (still supported)'], ['outOfSupport', 'Out of support'], ['other', 'Other / older']] },
];

function scaleOf(cut, max) {
  const clamp = (x) => Math.max(0, Math.min(max, Number(x) || 0));
  const points = [0, clamp(cut.criticalBelow), clamp(Math.max(cut.criticalBelow, cut.attentionBelow)),
    clamp(Math.max(cut.attentionBelow, cut.optimalFrom)), max];
  return ['Critical', 'Needs attention', 'Moderate', 'Optimal'].map((grade, i) => ({
    grade, width: `${((points[i + 1] - points[i]) / max) * 100}%`,
  }));
}

/** What counts as Critical … Optimal, per part and profile, and the grade colours. */
export default function StandardsPage({ devices, standards }) {
  const getToken = useSharePointToken();
  const { ask, dialog } = useConfirm();
  const [draft, setDraft] = useState(null);
  const [saving, setSaving] = useState(null);
  const [saveError, setSaveError] = useState('');

  const base = standards.standard;
  const current = draft ?? base;
  const edit = (change) => setDraft((was) => { const next = cloneStandard(was ?? base); change(next); return next; });

  const check = validateStandard(current);
  const errorAt = (path) => check.errors.find((e) => e.path === path)?.message;
  const lines = diffStandard(base, current);
  const preview = useMemo(() => previewChanges(devices, base, current), [devices, base, current]);
  const departments = useMemo(() => [...new Set(devices.map((d) => labelOf(d.department).toUpperCase()))]
    .filter((d) => d !== UNASSIGNED.toUpperCase()).sort(), [devices]);
  const readOnly = !standards.canEdit;

  const save = async (standard, summary, key) => {
    const yes = await ask({
      title: 'Save this standard?',
      body: `${summary}. ${preview.length ? `${preview.reduce((n, p) => n + p.count, 0)} part grades change across the fleet.` : 'No machine changes grade.'}`,
      confirmLabel: 'Save standard',
      cancelLabel: 'Keep editing',
    });
    if (!yes) return;
    setSaving(key);
    setSaveError('');
    try {
      const tokenRes = await getToken();
      await saveStandard({
        siteUrl: SITE, token: tokenRes.accessToken, standard, version: standards.version + 1,
        summary, savedBy: tokenRes.account?.username ?? '',
      });
      setDraft(null);
      standards.reload();
    } catch (failure) {
      setSaveError(failure.message);
    } finally {
      setSaving(null);
    }
  };

  return (
    <section className="sd">
      <header className="sd-head">
        <div>
          <h2 className="dm-title">Device standards</h2>
          <p className="dm-sub">What counts as Critical, Needs attention, Moderate and Optimal, for each part, against the work each profile does. Saving regrades every machine; nothing is re-scanned.</p>
          <p className="dm-sub">
            {standards.version ? `Version ${standards.version} · saved by ${standards.savedBy ?? 'unknown'}${standards.savedOn ? ` on ${formatMYT(standards.savedOn, 'date')}` : ''}` : 'The default standard — nothing saved yet'}
            {readOnly && ' · Only people with edit rights on the IT Device Standards list can change these. Ask IT.'}
          </p>
        </div>
        <div className="sd-actions">
          <Button variant="secondary" disabled={!draft || Boolean(saving)} onClick={() => setDraft(null)}>Discard changes</Button>
          <Button
            loading={saving === 'save'}
            disabled={readOnly || !draft || !check.ok || lines.length === 0}
            onClick={() => save(current, summaryOf(lines), 'save')}
          >
            {check.ok ? `Save standard${lines.length ? ` · ${lines.length} change${lines.length === 1 ? '' : 's'}` : ''}` : 'Fix the order to save'}
          </Button>
        </div>
      </header>

      {saveError && <ErrorBanner message={saveError} />}

      <div className="sd-body">
        <fieldset className="sd-main" disabled={readOnly || Boolean(saving)}>
          <legend className="sr-only">Standard</legend>
          <div className="sd-table">
            <div className="sd-row sd-row-head">
              <span>Part</span>
              {PROFILE_KEYS.map((key) => <span key={key}><strong>{PROFILE_SHORT[key]}</strong><small>{current.profiles[key].label}</small></span>)}
            </div>

            {NUMERIC.map(({ part, name, unit, max, note }) => (
              <div key={part} className="sd-part">
                <div className="sd-part-name"><strong>{name}</strong><small>{note}</small></div>
                {CUTS.map(({ cut, label, grade }) => (
                  <div key={cut} className="sd-row">
                    <span className="sd-row-label"><span className={`sd-dot g-${GRADE_SLUG[grade]}`} />{label}</span>
                    {PROFILE_KEYS.map((key) => {
                      const value = current.profiles[key][part][cut];
                      const changed = value !== base.profiles[key][part][cut];
                      return (
                        <label key={key} className="sd-num">
                          <span className="sr-only">{`${name}, ${PROFILE_SHORT[key]}, ${label}`}</span>
                          <input
                            type="number"
                            min="0"
                            className={changed ? 'sd-changed' : undefined}
                            value={Number.isFinite(value) ? value : ''}
                            onChange={(event) => edit((s) => {
                              s.profiles[key][part][cut] = event.target.value === '' ? null : Number(event.target.value);
                            })}
                          />
                          <span>{unit}</span>
                        </label>
                      );
                    })}
                  </div>
                ))}
                <div className="sd-row">
                  <span className="sd-row-label"><small>How the scale falls</small></span>
                  {PROFILE_KEYS.map((key) => (
                    <div key={key} className="sd-scale">
                      <span className="sd-scale-bar" aria-hidden="true">
                        {scaleOf(current.profiles[key][part], max).map((seg) => (
                          <span key={seg.grade} className={`g-${GRADE_SLUG[seg.grade]}`} style={{ width: seg.width }} />
                        ))}
                      </span>
                      {errorAt(`profiles.${key}.${part}`) && <small role="alert" className="sd-error">{errorAt(`profiles.${key}.${part}`)}</small>}
                    </div>
                  ))}
                </div>
              </div>
            ))}

            {CHOICES.map(({ part, name, note, rows }) => (
              <div key={part} className="sd-part">
                <div className="sd-part-name"><strong>{name}</strong><small>{note}</small></div>
                {rows.map(([kind, label]) => (
                  <div key={kind} className="sd-row">
                    <span className="sd-row-label">{label}</span>
                    {PROFILE_KEYS.map((key) => {
                      const value = current.profiles[key][part][kind];
                      return (
                        <label key={key} className="sd-choice">
                          <span className={`sd-dot g-${GRADE_SLUG[value] ?? 'unknown'}`} aria-hidden="true" />
                          <span className="sr-only">{`${name}, ${label}, ${PROFILE_SHORT[key]}`}</span>
                          <select
                            className={value !== base.profiles[key][part][kind] ? 'sd-changed' : undefined}
                            value={value}
                            onChange={(event) => edit((s) => { s.profiles[key][part][kind] = event.target.value; })}
                          >
                            {GRADES.map((g) => <option key={g} value={g}>{g}</option>)}
                          </select>
                        </label>
                      );
                    })}
                  </div>
                ))}
              </div>
            ))}
          </div>

          <div className="sd-card">
            <h3>Departments</h3>
            <p className="dm-sub">Which profile each department is judged against. Departments not listed, and machines with no department, use Desk.</p>
            <div className="sd-depts">
              {departments.map((department) => (
                <label key={department} className="sd-dept">
                  <span>{department}</span>
                  <select
                    value={current.departments[department] ?? 'desk'}
                    onChange={(event) => edit((s) => { s.departments[department] = event.target.value; })}
                  >
                    {PROFILE_KEYS.map((key) => <option key={key} value={key}>{PROFILE_SHORT[key]}</option>)}
                  </select>
                </label>
              ))}
            </div>
          </div>

          <div className="sd-card">
            <div className="sd-card-head">
              <div>
                <h3>Grade colours</h3>
                <p className="dm-sub">Used for every part chip, bar and badge in the device section. Colour only ever means a grade.</p>
              </div>
              <Button variant="secondary" onClick={() => edit((s) => { s.colors = { ...DEFAULT_COLORS }; })}>Reset to defaults</Button>
            </div>
            <div className="sd-colors">
              {[...GRADES, UNKNOWN].map((grade) => (
                <label key={grade} className="sd-color">
                  <input
                    type="color"
                    value={current.colors[grade]}
                    aria-label={`${grade} colour`}
                    onChange={(event) => edit((s) => { s.colors[grade] = event.target.value; })}
                  />
                  <span><strong>{grade}</strong><small>{current.colors[grade]}</small></span>
                  <span className="pc-chip" style={{ background: current.colors[grade], color: inkFor(current.colors[grade]) }}>RAM · {grade}</span>
                </label>
              ))}
            </div>
            {tooSimilar(current.colors) && <p role="alert" className="sd-warn">Two grade colours are hard to tell apart. You can still save, but people may misread a grade.</p>}
          </div>
        </fieldset>

        <aside className="sd-side">
          <div className="sd-card">
            <h3>If you save now</h3>
            {lines.length === 0 ? (
              <p className="dm-sub">No changes yet. Edit a number or a grade to see which machines would move.</p>
            ) : (
              <>
                <p className="dm-sub">{lines.length} change{lines.length === 1 ? '' : 's'} from {standards.version ? `version ${standards.version}` : 'the default'}.</p>
                <ul className="sd-preview">
                  {preview.map((p) => (
                    <li key={`${p.part}|${p.from}|${p.to}`}>
                      <strong>{p.part.toUpperCase()}</strong> · {p.count} machine{p.count === 1 ? '' : 's'}{' '}
                      <span className={`pc-chip g-${GRADE_SLUG[p.from]}`}>{p.from}</span> → <span className={`pc-chip g-${GRADE_SLUG[p.to]}`}>{p.to}</span>
                    </li>
                  ))}
                  {preview.length === 0 && <li>No machine changes grade.</li>}
                </ul>
              </>
            )}
          </div>

          <div className="sd-card">
            <h3>History</h3>
            {standards.history.length === 0 && <p className="dm-sub">Nothing saved yet — the default standard is in force.</p>}
            <ul className="sd-history">
              {standards.history.map((h) => (
                <li key={h.id}>
                  <span>
                    <strong>v{h.version ?? '?'}</strong> · {h.savedBy ?? 'unknown'}{h.savedOn ? ` · ${formatMYT(h.savedOn, 'date')}` : ''}
                    <small>{h.valid ? h.summary : 'Could not be read'}</small>
                  </span>
                  <Button
                    variant="secondary"
                    size="sm"
                    loading={saving === `restore-${h.id}`}
                    disabled={readOnly || !h.valid || Boolean(saving) || h.version === standards.version}
                    onClick={() => save(h.standard, `Restored version ${h.version}`, `restore-${h.id}`)}
                  >
                    Restore
                  </Button>
                </li>
              ))}
            </ul>
          </div>
        </aside>
      </div>
      {dialog}
    </section>
  );
}
```

`save` closes over `preview`, which describes the DRAFT. For a restore, the dialog's count still describes the draft. That is acceptable for a restore, which replaces the standard wholesale, but say it in the report.

- [ ] **Step 2: Wire the tab**

In `src/pages/DevicesPage.jsx`:
- Import `StandardsPage from '../features/devices/ui/StandardsPage'`.
- Add `['standards', 'Standards'],` as the LAST tab.
- Render `{view === 'standards' && <StandardsPage devices={saved} standards={standards} />}`.

- [ ] **Step 3: Style it**

Append to `src/styles/devices.css`:

```css
/* --- Standards tab ------------------------------------------------------ */
.sd { display: flex; flex-direction: column; gap: 16px; }
.sd-head { display: flex; flex-wrap: wrap; justify-content: space-between; align-items: flex-end; gap: 16px; }
.sd-head .dm-sub { max-width: 760px; margin: 4px 0 0; }
.sd-actions { display: flex; gap: 8px; flex-wrap: wrap; }
.sd-body { display: flex; flex-wrap: wrap; gap: 16px; align-items: flex-start; }
.sd-main { flex: 999 1 640px; min-width: 0; border: 0; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 16px; }
.sd-side { flex: 1 1 300px; min-width: 0; display: flex; flex-direction: column; gap: 16px; }
.sd-table, .sd-card { background: var(--it-panel); border: 1px solid var(--it-line); border-radius: 16px; }
.sd-table { overflow-x: auto; }
.sd-card { padding: 16px 18px; display: flex; flex-direction: column; gap: 10px; }
.sd-card h3 { margin: 0; font-size: 17px; }
.sd-card-head { display: flex; flex-wrap: wrap; justify-content: space-between; gap: 10px; align-items: flex-end; }
.sd-row { display: grid; grid-template-columns: 220px repeat(3, minmax(160px, 1fr)); gap: 0 14px; align-items: center; padding: 4px 18px; min-width: 760px; }
.sd-row-head { padding: 14px 18px; border-bottom: 1px solid var(--it-line); font-size: 12px; color: var(--it-ink-soft); }
.sd-row-head strong { display: block; font-size: 15px; color: var(--it-ink); }
.sd-part { padding: 12px 0; border-bottom: 1px solid var(--it-line); }
.sd-part-name { padding: 0 18px 6px; display: flex; flex-direction: column; gap: 2px; }
.sd-part-name small, .sd-row-label small { color: var(--it-ink-soft); font-size: 13px; }
.sd-row-label { display: flex; align-items: center; gap: 8px; font-size: 14px; }
.sd-dot { flex: none; width: 12px; height: 12px; border-radius: 50%; }
.sd-num, .sd-choice { display: flex; align-items: center; gap: 8px; }
.sd-num input, .sd-choice select, .sd-dept select {
  flex: 1; min-width: 0; min-height: 44px; padding: 0 10px; font: inherit; font-size: 16px;
  border: 1px solid var(--it-line); border-radius: 10px; background: var(--it-panel); color: var(--it-ink);
}
.sd-num span { font-size: 13px; color: var(--it-ink-soft); }
.sd-changed { border-color: var(--it-brand) !important; background: var(--it-brand-wash) !important; }
.sd-scale { display: flex; flex-direction: column; gap: 4px; }
.sd-scale-bar { display: flex; height: 10px; border-radius: 999px; overflow: hidden; background: var(--it-line); }
.sd-scale-bar > span { display: block; height: 100%; }
.sd-error { color: var(--it-danger); font-weight: 600; }
.sd-depts { display: grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap: 10px; }
.sd-dept { display: flex; align-items: center; gap: 10px; }
.sd-dept > span { flex: 0 0 120px; font-weight: 600; font-family: ui-monospace, 'SFMono-Regular', Menlo, monospace; }
.sd-colors { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 10px; }
.sd-color { display: flex; align-items: center; gap: 12px; padding: 8px 10px; border: 1px solid var(--it-line); border-radius: 12px; }
.sd-color input { width: 44px; height: 44px; border: 0; padding: 0; background: transparent; cursor: pointer; }
.sd-color > span:nth-child(2) { flex: 1; display: flex; flex-direction: column; }
.sd-color small { color: var(--it-ink-soft); font-family: ui-monospace, 'SFMono-Regular', Menlo, monospace; }
.sd-warn { margin: 0; padding: 8px 12px; border-radius: 10px; background: var(--it-accent-wash); color: var(--it-ink); font-weight: 600; }
.sd-preview, .sd-history { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 8px; font-size: 14px; }
.sd-history li { display: flex; justify-content: space-between; align-items: flex-start; gap: 10px; padding-top: 8px; border-top: 1px solid var(--it-line); }
.sd-history small { display: block; color: var(--it-ink-soft); }
```

- [ ] **Step 4: Check**

Run: `npx eslint src`, `npm run build`, `npm test`
Expected: only the baseline lint errors, a passing build, and green tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/ui/StandardsPage.jsx src/pages/DevicesPage.jsx src/styles/devices.css
git commit -m "Add the Standards tab: edit each part's cut-offs and grades per profile, the department map and the grade colours, with a preview and history

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 10: Write it down, and check the whole thing

**Files:**
- Modify: `AGENTS.md`

- [ ] **Step 1: Update `AGENTS.md`**

In ROUTES, add `standards` to the `/devices` row's list of views, and mention that it is where IT sets the grading standard and colours.

In WHERE TO LOOK, add these rows:

```
| What counts as Critical … Optimal, per part and profile | `src/features/devices/standards/defaultStandard.js`, the Standards tab, the `IT Device Standards` list |
| How one part of a machine is graded | `src/features/devices/derive/partGrades.js` |
| Whether a saved standard is usable, and what a save changed | `standards/validateStandard.js`, `standards/diffStandard.js`, `standards/previewChanges.js` |
| Grade colours | `standards/gradeColors.js`; `--grade-*` variables on `.dv-graded` |
```

Replace the paragraph that begins **"A machine is judged against the desk it sits on, not against one fleet-wide bar."** and the one after it ("The whole persona layer is computed on read …") with:

```markdown
**A machine is graded part by part, against the desk it sits on, by a standard
IT sets.** Five parts -- CPU, RAM, Storage, Graphics, Windows -- each get their
own grade (Critical, Needs attention, Moderate, Optimal, or Unknown when the
scan did not say), and there is deliberately NO overall grade: one word for the
whole machine hid which part to fix. Each workload profile (Engineering, Desk,
Field) has its own cut-offs, and the department → profile map is part of the
same standard. A value `< criticalBelow` is Critical, `< attentionBelow` Needs
attention, `< optimalFrom` Moderate, otherwise Optimal; storage takes the worse
of disk type and size; obsolete CPU families are always Critical.

The standard is DATA: versions in the `IT Device Standards` list, append-only,
the newest valid one in force, edited on the Standards tab by whoever SharePoint
lets add items to that list. A broken newest version is skipped and named, never
silently. Grades are computed on read (`useDevices` → `regrade`) and nothing of
them is stored, so a saved standard regrades the fleet on the next render with
no re-scan. Colour in the device section only ever means a grade, and comes from
the standard's five colours through `--grade-*` CSS variables -- never a
hard-coded red.
```

- [ ] **Step 2: Full check**

Run: `npm test`, `npx eslint src`, `npm run build`
Expected: green tests, only the baseline lint errors, and a passing build.

Run: `grep -rn "fitStatus\|fitReasons\|deviceFit\|FIT_ORDER\|'Needs Attention'" src`
Expected: no hits.

Then start the production preview with the preview tool (`pmw-it-preview` in `.claude/launch.json`, port 4173). Open `/devices?view=standards` and check that the sign-in gate renders with no console errors. A signed-in check against real rows is left to the user.

- [ ] **Step 3: Commit**

```bash
git add AGENTS.md
git commit -m "Document part grades, the device standard and grade colours

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

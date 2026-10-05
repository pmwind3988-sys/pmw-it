# Device Map and Machine Lifecycle Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `/devices` becomes a map you walk into (location → department → machine). Every machine has a status (In use / In repair / Spare / Retired), a serial-based identity and an owner history, so a new laptop, a hand-me-down and a retirement are each recorded without losing the old record.

**Architecture:** The rules are pure modules under `src/features/devices/` (`lifecycle/`, `map/`, `derive/deriveLocation.js`). Each one is tested without React or SharePoint. SharePoint gains four columns on `IT Device List`, one on `IT Device Changes`, and a new `IT Device Assignments` list. An import plans everything at once in `planImport` (it replaces `planSync`). A lifecycle action re-reads the machine, plans with `planLifecycle`, then writes in this order: device row → stints → change log. The screens are a new Map tab (the new default), an extended machine page, a replacement prompt in the import review, and status/location filters on the register.

**Tech Stack:** React 19, React Router, Vite 8, Vitest 3, SharePoint REST through `spFetch`.

**Spec:** `docs/superpowers/specs/2026-10-05-device-map-and-lifecycle-design.md`. Read it before starting any task. The mockup is at https://claude.ai/artifact/PuQ6YuUwvwWG8aumwkvBjD.

## Global Constraints

- Work in the worktree `C:\Users\User\pmw-it\.claude\worktrees\devices-map`, branch `devices-map-lifecycle`. Never touch the main checkout. Another session is editing checklist-link files there.
- Do not modify any checklist-link file: `server/`, `api/`, `src/public/`, `src/features/forms/links/`, `src/pages/ChecklistLinks*`, or `src/App.jsx`. The `/devices/:id` route already exists.
- **Statuses** are exactly `In use`, `In repair`, `Spare`, `Retired`. A blank status reads as `In use`.
- **End reasons** are exactly `Reassigned`, `Replaced`, `To stash`, `To repair`, `Retired`.
- The new list is titled exactly `IT Device Assignments`.
- **Built-in locations** are `F1`, `F3`, `PML`. Locations are stored as upper-case codes in a TEXT column, never a choice column.
- **SharePoint columns:**
  - Create each through the existing `provisionSchema` declarations only. That takes care of SP.Field concrete types, internal names, `DisplayFormat: 1`, and `RichText: false`.
  - Never send an empty multi-choice.
  - Partial writes send only the columns being changed (MERGE). Never use `toListItem` for a partial write.
- **Write order, always:** device row first, then stints, then the change log.
- **Date-only inputs** are parsed with `parseFormDate` (local noon), never with `new Date('YYYY-MM-DD')`.
- **React:**
  - No `navigate()` inside `useEffect`.
  - No state reset from an effect. Clamp or compare during render.
  - No ref writes during render.
  - No helper exported next to a component in the same file.
- **Every button shows when it is working.** Any button that starts a SharePoint read or write shows the spinner from the press until the work settles. Use `<Button loading>`, or `<Spinner />` inside a raw `<button>` (Task 9a). The label stays as it is, a second press is blocked, and only the pressed button spins. Buttons that start nothing (Cancel, close, a tab) do not get it.
- **Icons:** add any new glyph to `src/components/ui/Icons.jsx`, and import every icon a component uses.
- **Copy rules:**
  - No emoji.
  - Inputs in CSS must be ≥16px font-size.
  - Touch targets ≥44px.
- **`npm run lint`** may report only its one existing error (`ThemeContext.jsx`). Nothing new.
- **Commits:** end each with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- **Tests:** `npx vitest run <path>` for one file, `npm test` for all.

## File map

| File | Status | Responsibility |
|---|---|---|
| `src/features/devices/parse/placeholders.js` | modify | `normaliseSerial`, `isPlaceholderSerial`, `cleanSerial` |
| `src/features/devices/parse/labels.js` | modify | `Serial Number` label |
| `src/features/devices/map/locations.js` | create | built-in + in-use locations |
| `src/features/devices/derive/deriveLocation.js` | create | location out of the file-name bracket |
| `src/features/devices/derive/deriveIdentity.js` | modify | location before department; two new department aliases |
| `src/features/devices/derive/deriveDevice.js` | modify | `serialNumber`, `knownLocations` |
| `src/features/devices/derive/persona.js` | modify | `STOCKYARD`, `GUARDHOUSE` → DESK |
| `src/features/devices/importFiles.js` | modify | pass `knownLocations` through |
| `src/features/devices/lifecycle/status.js` | create | statuses, `statusOf`, `inFleet` |
| `src/features/devices/lifecycle/stints.js` | create | open/legacy stint, timeline |
| `src/features/devices/lifecycle/planLifecycle.js` | create | one action → fields, changes, close, open |
| `src/features/devices/lifecycle/matchIncoming.js` | create | serial-then-name matching |
| `src/features/devices/lifecycle/replacements.js` | create | "is this a replacement?" prompts |
| `src/features/devices/lifecycle/specHistory.js` | create | change rows → dated spec history |
| `src/features/devices/sharepoint/deviceSchema.js` | modify | 4 columns, `DeviceId`, tracked fields, `createdOn` |
| `src/features/devices/sharepoint/assignmentSchema.js` | create | the stints list schema |
| `src/features/devices/sharepoint/provisionLists.js` | modify | third list |
| `src/features/devices/sharepoint/readDevices.js` | modify | `readDevice(id)` |
| `src/features/devices/sharepoint/readAssignments.js` | create | stints read |
| `src/features/devices/sharepoint/readHistory.js` | create | change rows for one machine |
| `src/features/devices/sharepoint/writeLifecycle.js` | create | `performLifecycle`, `lifecycleItem`, `stintWrites`, `writeStints` |
| `src/features/devices/sharepoint/updateDevice.js` | modify | `location` editable; export `logChanges` with `DeviceId` |
| `src/features/devices/sharepoint/syncDevices.js` | modify | `planImport` replaces `planSync` |
| `src/features/devices/reviewIssues.js` | modify | `sortForReview(devices, first)` |
| `src/features/devices/deviceFilters.js` | modify | `status`, `location` matchers |
| `src/features/devices/map/zones.js` | create | tile summaries for every level |
| `src/features/devices/map/mapLayout.js` | create | grid cells, arrow-key neighbour, connectors |
| `src/features/devices/map/mapLinks.js` | create | map addresses |
| `src/features/devices/useDeviceHistory.js` | create | stints + changes for one machine |
| `src/features/devices/ui/ReplacementPrompt.jsx` | create | the three-way choice |
| `src/features/devices/ui/ReviewGrid.jsx` | modify | prompt + notice rows |
| `src/features/devices/ui/MapTile.jsx`, `MapGrid.jsx`, `WorldMap.jsx`, `LocationView.jsx`, `DepartmentView.jsx`, `MachineCard.jsx`, `DeviceMap.jsx` | create | the map screens |
| `src/features/devices/ui/LifecycleActions.jsx`, `AssignDialog.jsx`, `OwnerHistory.jsx`, `SpecHistory.jsx` | create | the machine page additions |
| `src/features/devices/ui/DeviceTable.jsx`, `src/features/devices/fieldGroups.js` | modify | new columns lead / grouped |
| `src/pages/DevicesPage.jsx`, `src/pages/DeviceDetailPage.jsx` | modify | wire it all |
| `src/components/ui/Spinner.jsx` | create | the one busy mark |
| `src/components/ui/Button.jsx`, `src/components/ui/Surfaces.jsx` | modify | `loading` on Button, `busy` on ErrorBanner's Retry |
| `src/components/ui/Icons.jsx` | modify | `Monitor`, `Archive`, `Tombstone`, `Map`, `Wrench` |
| `src/styles/devices.css` | modify | map, cards, prompt, history styles |
| `AGENTS.md` | modify | conventions, where-to-look, routes |

---

### Task 1: Read the serial number from the scan

**Files:**
- Modify: `src/features/devices/parse/placeholders.js`
- Modify: `src/features/devices/parse/labels.js`
- Modify: `src/features/devices/derive/deriveDevice.js`
- Test: `src/features/devices/parse/serial.test.js` (create), `src/features/devices/parse/labels.test.js`, `src/features/devices/derive/deriveDevice.test.js`

**Interfaces:**
- Produces:
  - `normaliseSerial(value) → string` (upper case, no whitespace or invisibles)
  - `isPlaceholderSerial(value) → boolean`
  - `cleanSerial(value) → string|null`
  - `device.serialNumber: string|null` on every derived device

- [ ] **Step 1: Write the failing tests**

Create `src/features/devices/parse/serial.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { cleanSerial, isPlaceholderSerial, normaliseSerial } from './placeholders.js';

describe('normaliseSerial', () => {
  it('upper-cases and removes every space', () => {
    expect(normaliseSerial(' 5cg8241 kqz ')).toBe('5CG8241KQZ');
  });

  it('removes invisible characters too', () => {
    expect(normaliseSerial('5CG\u00a08241\u200bKQZ')).toBe('5CG8241KQZ');
  });
});

describe('cleanSerial', () => {
  it('keeps a real serial', () => {
    expect(cleanSerial('MXL7180Q2B')).toBe('MXL7180Q2B');
  });

  it.each([
    'System Serial Number', 'Chassis Serial Number', 'Default string',
    'To be filled by O.E.M.', '0', '00000000', 'XXXXXXXX', '', '   ', '\u00a0', null, undefined,
  ])('treats %j as no serial', (value) => {
    expect(cleanSerial(value)).toBeNull();
    expect(isPlaceholderSerial(value)).toBe(true);
  });

  it('does not treat a serial that merely contains a repeat as junk', () => {
    expect(cleanSerial('5CG000001')).toBe('5CG000001');
  });
});
```

In `src/features/devices/parse/labels.test.js`, change `toHaveLength(21)` to `toHaveLength(22)`, and add inside the same `describe('KNOWN_LABELS'…)`:

```js
  it('knows the serial number line', () => {
    expect(KNOWN_LABELS).toContain('Serial Number');
  });
```

In `src/features/devices/derive/deriveDevice.test.js`, append:

```js
describe('deriveDevice — serial number', () => {
  it('reads the serial line the scan script writes', () => {
    const report = load('ASHRAF-PC_.txt');
    const device = deriveDevice({ ...report, text: `Serial Number: 5cg8241kqz\n${report.text}` });
    expect(device.serialNumber).toBe('5CG8241KQZ');
  });

  it('has no serial on a report written before the line existed', () => {
    expect(deriveDevice(load('ASHRAF-PC_.txt')).serialNumber).toBeNull();
  });

  it('drops a placeholder serial from a home-built desktop', () => {
    const report = load('ASHRAF-PC_.txt');
    const device = deriveDevice({ ...report, text: `Serial Number: Default string\n${report.text}` });
    expect(device.serialNumber).toBeNull();
  });
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `npx vitest run src/features/devices/parse src/features/devices/derive/deriveDevice.test.js`
Expected: FAIL. `cleanSerial` is not exported, the label count is 21, and `serialNumber` is undefined.

- [ ] **Step 3: Implement**

Append to `src/features/devices/parse/placeholders.js`:

```js
/**
 * What a BIOS writes when nobody flashed a real serial in. A home-built
 * desktop reports one of these, and every such desktop reports the SAME one,
 * so matching on it would merge every unbranded PC in the company into one row.
 */
const PLACEHOLDER_SERIALS = new Set([
  'system serial number',
  'chassis serial number',
  'default string',
  'to be filled by o.e.m.',
  'not specified',
  '0',
]);

/** Serials compare upper-cased with every space gone: `cn0abc 123` is `CN0ABC123`. */
export function normaliseSerial(value) {
  return String(value ?? '').replace(INVISIBLE, '').replace(/\s+/g, '').toUpperCase();
}

export function isPlaceholderSerial(value) {
  if (isPlaceholder(value)) return true;
  const spaced = String(value).replace(INVISIBLE, ' ').trim().toLowerCase();
  if (PLACEHOLDER_SERIALS.has(spaced)) return true;
  // `00000000`, `XXXXXXXX`: one character repeated is a filler, not a serial.
  return /^(.)\1*$/.test(normaliseSerial(value));
}

/** The serial as stored, or null when the scan had nothing usable. */
export function cleanSerial(value) {
  return isPlaceholderSerial(value) ? null : normaliseSerial(value);
}
```

In `src/features/devices/parse/labels.js`, append `'Serial Number',` as the last entry of `KNOWN_LABELS`, with this comment line above it:

```js
  // Added 2026-10-05 so a machine keeps its history through a rename. The scan
  // script writes it as `Serial Number: $((Get-CimInstance Win32_BIOS).SerialNumber)`.
  'Serial Number',
```

In `src/features/devices/derive/deriveDevice.js`:
- Change the import to `import { cleanValue, cleanSerial } from '../parse/placeholders.js';`
- Inside `base`, after `anydeskId`, add:

```js
    serialNumber: cleanSerial(fields['Serial Number']?.[0]),
```

- [ ] **Step 4: Run them to see them pass**

Run: `npx vitest run src/features/devices/parse src/features/devices/derive`
Expected: PASS. If the inline-value test fails because `parseReport` does not return an inline value for a label, check `parseReport.test.js` for how `Antivirus status: …` is read. It is inline in every fixture, so `Serial Number:` takes the same path.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/parse src/features/devices/derive/deriveDevice.js src/features/devices/derive/deriveDevice.test.js
git commit -m "Read the machine's serial number from the scan, and treat BIOS filler as no serial

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Read the location from the file name

**Files:**
- Create: `src/features/devices/map/locations.js`, `src/features/devices/map/locations.test.js`
- Create: `src/features/devices/derive/deriveLocation.js`, `src/features/devices/derive/deriveLocation.test.js`
- Modify: `src/features/devices/derive/deriveIdentity.js`, `src/features/devices/derive/deriveDevice.js`, `src/features/devices/derive/persona.js`, `src/features/devices/importFiles.js`
- Modify tests: `src/features/devices/derive/deriveIdentity.test.js:36-40`, `src/features/devices/derive/deriveDevice.test.js:67-68`

**Interfaces:**
- Produces:
  - `BUILT_IN_LOCATIONS: string[]`
  - `cleanLocation(typed) → string|null` (trimmed, upper case)
  - `locationsIn(devices) → string[]` (built-in first, then in-use, no duplicates)
  - `splitLocation(bracket, known = BUILT_IN_LOCATIONS) → { location: string|null, rest: string|null }`
  - `device.location: string|null`
  - `deriveDevice({ text, fileName, lastModified, knownLocations })`
  - `importFiles(files, { knownLocations } = {})`

- [ ] **Step 1: Write the failing tests**

Create `src/features/devices/map/locations.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { BUILT_IN_LOCATIONS, cleanLocation, locationsIn } from './locations.js';

describe('cleanLocation', () => {
  it('upper-cases and trims', () => expect(cleanLocation('  f1 ')).toBe('F1'));
  it('reads a blank as no location', () => {
    expect(cleanLocation('')).toBeNull();
    expect(cleanLocation(null)).toBeNull();
  });
});

describe('locationsIn', () => {
  it('starts with the built-in locations', () => {
    expect(locationsIn([])).toEqual(BUILT_IN_LOCATIONS);
    expect(BUILT_IN_LOCATIONS).toEqual(['F1', 'F3', 'PML']);
  });

  it('adds a location the register is using, once, whatever its case', () => {
    expect(locationsIn([{ location: 'hq' }, { location: 'HQ' }, { location: 'f1' }, { location: '' }]))
      .toEqual(['F1', 'F3', 'PML', 'HQ']);
  });
});
```

Create `src/features/devices/derive/deriveLocation.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { splitLocation } from './deriveLocation.js';

describe('splitLocation', () => {
  it('takes a known code off the front of the bracket', () => {
    expect(splitLocation('F1 ENGINEERING')).toEqual({ location: 'F1', rest: 'ENGINEERING' });
    expect(splitLocation('pml finance evonne')).toEqual({ location: 'PML', rest: 'finance evonne' });
  });

  it('leaves a bracket with no location exactly as it was', () => {
    expect(splitLocation('ENGINEERING')).toEqual({ location: null, rest: 'ENGINEERING' });
    expect(splitLocation('QAQC FAIRUS')).toEqual({ location: null, rest: 'QAQC FAIRUS' });
  });

  it('reads a bracket that is only a location', () => {
    expect(splitLocation('F3')).toEqual({ location: 'F3', rest: null });
  });

  it('splits a code glued to the end of the department', () => {
    expect(splitLocation('STOCKYARDF1')).toEqual({ location: 'F1', rest: 'STOCKYARD' });
  });

  it('splits the two-word guardhouse name', () => {
    expect(splitLocation('PML GUARDHOUSE')).toEqual({ location: 'PML', rest: 'GUARDHOUSE' });
  });

  it('does not cut a short word that only ends like a code', () => {
    expect(splitLocation('XF1')).toEqual({ location: null, rest: 'XF1' });
  });

  it('recognises a location added in the portal', () => {
    expect(splitLocation('HQ ADMIN', ['F1', 'HQ'])).toEqual({ location: 'HQ', rest: 'ADMIN' });
  });

  it('has nothing to split without a bracket', () => {
    expect(splitLocation(null)).toEqual({ location: null, rest: null });
  });
});
```

In `src/features/devices/derive/deriveIdentity.test.js`, replace the body of `it('keeps a two-word department whole'…)` (lines 36-40) with:

```js
  it('reads PML GUARDHOUSE as a location and a department', () => {
    const result = deriveIdentity(withFields({ 'Computer Name': ['PMWP001'] }),
      '[PML GUARDHOUSE] PMWP001_.txt');
    expect(result.location).toBe('PML');
    expect(result.department).toBe('GUARDHOUSE');
  });

  it('reads the location in front of the department', () => {
    const result = deriveIdentity(withFields({ 'Computer Name': ['AMIR-HP'] }),
      '[F1 ENGINEERING AMIR] AMIR-HP_.txt');
    expect(result.location).toBe('F1');
    expect(result.department).toBe('ENGINEERING');
    expect(result.owner).toBe('Amir');
  });
```

In `src/features/devices/derive/deriveDevice.test.js` at lines 67-68, replace `expect(device.department).toBe('STOCKYARDF1');` with:

```js
    expect(device.location).toBe('F1');
    expect(device.department).toBe('STOCKYARD');
```

- [ ] **Step 2: Run them to see them fail**

Run: `npx vitest run src/features/devices/map src/features/devices/derive`
Expected: FAIL. The modules don't exist, and `location` is undefined.

- [ ] **Step 3: Implement**

Create `src/features/devices/map/locations.js`:

```js
/**
 * Where a machine physically is: F1, F3, PML and whatever else IT types.
 *
 * Stored as an upper-case CODE in a text column, never a choice column: a new
 * site must not need a SharePoint column change before a machine can be put
 * there. Which locations exist is read the way `assets/categories.js` reads
 * categories -- the built-in ones, then every one a row is actually using --
 * so there is no second list to disagree with the rows.
 */
export const BUILT_IN_LOCATIONS = ['F1', 'F3', 'PML'];

export function cleanLocation(typed) {
  const code = String(typed ?? '').trim().replace(/\s+/g, ' ').toUpperCase();
  return code || null;
}

export function locationsIn(devices = []) {
  const list = [...BUILT_IN_LOCATIONS];
  const seen = new Set(list);
  for (const device of devices) {
    const code = cleanLocation(device?.location);
    if (!code || seen.has(code)) continue;
    seen.add(code);
    list.push(code);
  }
  return list;
}
```

Create `src/features/devices/derive/deriveLocation.js`:

```js
import { BUILT_IN_LOCATIONS, cleanLocation } from '../map/locations.js';

/**
 * The location IT writes first inside the file name's bracket:
 * `[F1 ENGINEERING] AMIR-HP.txt`. Anything that is not a known code is left
 * alone, so `[ENGINEERING] AMIR-HP.txt` reads exactly as it always did.
 *
 * Two names predate the convention and carry the site glued on or in front:
 * `STOCKYARDF1` and `PML GUARDHOUSE`. A glued code is split off only when it is
 * a KNOWN code with at least three letters in front of it, so a word that
 * merely ends in something code-shaped is not cut.
 */
export function splitLocation(bracket, known = BUILT_IN_LOCATIONS) {
  if (!bracket || !bracket.trim()) return { location: null, rest: null };

  const codes = [...new Set(known.map(cleanLocation).filter(Boolean))]
    .sort((a, b) => b.length - a.length);
  const words = bracket.trim().split(/\s+/);
  const first = words[0].toUpperCase();
  const tail = words.slice(1).join(' ') || null;

  if (codes.includes(first)) return { location: first, rest: tail };

  for (const code of codes) {
    if (first.length < code.length + 3 || !first.endsWith(code)) continue;
    const head = words[0].slice(0, -code.length);
    if (!/[A-Z]$/i.test(head)) continue;
    return { location: code, rest: [head, tail].filter(Boolean).join(' ') };
  }

  return { location: null, rest: bracket.trim() };
}
```

In `src/features/devices/derive/deriveIdentity.js`:
- Add the import `import { splitLocation } from './deriveLocation.js';`
- In `KNOWN_DEPARTMENTS`, add `'STOCKYARD', 'GUARDHOUSE',` after `'PURCHASING',`. Keep the old `'STOCKYARDF1'` and `'PML GUARDHOUSE'` entries: they are harmless now that the location comes off first.
- Replace `deriveIdentity` with:

```js
export function deriveIdentity(fields, fileName, { knownLocations } = {}) {
  const { bracket, stem } = parseFileName(fileName);
  const { location, rest } = splitLocation(bracket, knownLocations);
  const { department, person } = splitBracket(rest);

  const fromField = fields['Computer Name']?.length ? cleanValue(fields['Computer Name'][0]) : null;

  return {
    computerName: fromField ?? (stem || null),
    location,
    department,
    ...resolveOwner(fields, person),
    ...resolveDeviceType(fields),
  };
}
```

In `src/features/devices/derive/deriveDevice.js`:
- Change the signature to `export function deriveDevice({ text, fileName, lastModified, knownLocations })`.
- Change the identity line to `const identity = deriveIdentity(fields, fileName, { knownLocations });`

In `src/features/devices/derive/persona.js`, add two entries next to `STOCKYARDF1: DESK,`:

```js
  STOCKYARD: DESK,
  GUARDHOUSE: DESK,
```

In `src/features/devices/importFiles.js`:
- Change the signature to `export async function importFiles(files, { knownLocations } = {}) {`
- Change the push to `devices.push(deriveDevice({ text, fileName: file.name, lastModified: file.lastModified, knownLocations }));`

- [ ] **Step 4: Run them to see them pass, then run the whole device suite**

Run: `npx vitest run src/features/devices`
Expected: PASS. A persona test that pinned `STOCKYARDF1` still passes, because the old key is kept.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/map src/features/devices/derive src/features/devices/importFiles.js
git commit -m "Read a machine's location from the scan file name: [F1 ENGINEERING], and split STOCKYARDF1 and PML GUARDHOUSE

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: Statuses, and the new device columns

**Files:**
- Create: `src/features/devices/lifecycle/status.js`, `src/features/devices/lifecycle/status.test.js`
- Modify: `src/features/devices/sharepoint/deviceSchema.js`, `src/features/devices/sharepoint/deviceSchema.test.js`
- Modify: `src/features/devices/sharepoint/updateDevice.js`, `src/features/devices/sharepoint/updateDevice.test.js:17`

**Interfaces:**
- Produces:
  - `STATUSES`, `IN_USE`, `IN_REPAIR`, `SPARE`, `RETIRED`
  - `statusOf(device) → string`
  - `inFleet(device) → boolean`
  - `DEVICE_COLUMNS` gains `Location`, `SerialNumber`, `Status`, `StatusChangedOn`. `CHANGE_COLUMNS` gains `DeviceId`.
  - `fromListItem(row).createdOn: number|null`
  - `TRACKED_FIELDS` gains `computerName`, `location`, `serialNumber`, `status`
  - `EDITABLE_FIELDS` = `['owner', 'department', 'deviceType', 'location']`
  - `logChanges(siteUrl, token, digest, device, changes, changedBy)` exported, writing `DeviceId: device.id`

- [ ] **Step 1: Write the failing tests**

Create `src/features/devices/lifecycle/status.test.js`:

```js
import { describe, it, expect } from 'vitest';
import {
  STATUSES, statusOf, inFleet, IN_USE, IN_REPAIR, SPARE, RETIRED,
} from './status.js';

describe('statusOf', () => {
  it('reads a blank status as In use, so no row needs migrating', () => {
    expect(statusOf({})).toBe(IN_USE);
    expect(statusOf({ status: '' })).toBe(IN_USE);
    expect(statusOf({ status: 'Nonsense' })).toBe(IN_USE);
  });

  it('keeps a real status', () => {
    expect(statusOf({ status: 'Retired' })).toBe(RETIRED);
  });

  it('lists exactly four statuses', () => {
    expect(STATUSES).toEqual(['In use', 'In repair', 'Spare', 'Retired']);
  });
});

describe('inFleet', () => {
  it('counts machines in use and in repair, not spare or retired ones', () => {
    expect(inFleet({ status: IN_USE })).toBe(true);
    expect(inFleet({ status: IN_REPAIR })).toBe(true);
    expect(inFleet({})).toBe(true);
    expect(inFleet({ status: SPARE })).toBe(false);
    expect(inFleet({ status: RETIRED })).toBe(false);
  });
});
```

Append to `src/features/devices/sharepoint/deviceSchema.test.js`:

```js
describe('lifecycle columns', () => {
  const names = DEVICE_COLUMNS.map((column) => column.StaticName);

  it('adds location, serial, status and when the status changed', () => {
    expect(names).toEqual(expect.arrayContaining(['Location', 'SerialNumber', 'Status', 'StatusChangedOn']));
  });

  it('stores location as text so a new site needs no column change', () => {
    expect(DEVICE_COLUMNS.find((c) => c.StaticName === 'Location').kind).toBe('text');
  });

  it('offers exactly the four statuses', () => {
    expect(DEVICE_COLUMNS.find((c) => c.StaticName === 'Status').choices)
      .toEqual(['In use', 'In repair', 'Spare', 'Retired']);
  });

  it('ties change rows to the machine, not only its name', () => {
    expect(CHANGE_COLUMNS.find((c) => c.StaticName === 'DeviceId').kind).toBe('number');
  });

  it('logs renames, location, serial and status changes', () => {
    expect(TRACKED_FIELDS).toEqual(expect.arrayContaining(['computerName', 'location', 'serialNumber', 'status']));
  });

  it('reads when the row was first created', () => {
    const record = fromListItem({ Id: 3, Title: 'PC1', Created: '2026-08-21T02:00:00Z' });
    expect(record.createdOn).toBe(Date.parse('2026-08-21T02:00:00Z'));
    expect(fromListItem({ Id: 3, Title: 'PC1' }).createdOn).toBeNull();
  });

  it('round-trips a status and a location', () => {
    const item = toListItem({ computerName: 'PC1', status: 'Spare', location: 'F1' });
    expect(item.Status).toBe('Spare');
    expect(item.Location).toBe('F1');
  });
});
```

In `src/features/devices/sharepoint/updateDevice.test.js` line 17, change the expectation to:

```js
    expect(EDITABLE_FIELDS).toEqual(['owner', 'department', 'deviceType', 'location']);
```

Append to the same file, inside `describe('updateDevice', …)`:

```js
  it('ties every change row to the machine id', async () => {
    const sp = fakeSharePoint();
    vi.stubGlobal('fetch', sp.fetch);

    await updateDevice({
      siteUrl: SITE, token: 't', existing: row(), edits: { location: 'F3' }, changedBy: 'me',
    });

    const logged = writes(sp.calls).find((c) => c.url.includes('IT%20Device%20Changes'));
    expect(logged.body.DeviceId).toBe(7);
    expect(logged.body.FieldName).toBe('location');
  });
```

- [ ] **Step 2: Run them to see them fail**

Run: `npx vitest run src/features/devices/lifecycle src/features/devices/sharepoint/deviceSchema.test.js src/features/devices/sharepoint/updateDevice.test.js`
Expected: FAIL.

- [ ] **Step 3: Implement**

Create `src/features/devices/lifecycle/status.js`:

```js
/**
 * Where a machine is in its life. A blank status reads as In use, which is
 * what every row saved before statuses existed actually is -- so the column
 * needed no migration.
 *
 * Only In use and In repair count towards the fleet's figures. A retired
 * laptop reported as a "critical risk" would be a figure lying about the
 * machines people are working on.
 */
export const IN_USE = 'In use';
export const IN_REPAIR = 'In repair';
export const SPARE = 'Spare';
export const RETIRED = 'Retired';

export const STATUSES = [IN_USE, IN_REPAIR, SPARE, RETIRED];

export function statusOf(device) {
  return STATUSES.includes(device?.status) ? device.status : IN_USE;
}

export function inFleet(device) {
  const status = statusOf(device);
  return status === IN_USE || status === IN_REPAIR;
}
```

In `src/features/devices/sharepoint/deviceSchema.js`:
- Add `import { STATUSES } from '../lifecycle/status.js';`
- In `DEVICE_COLUMNS`, after `text('Department', 'Department'),`, add `text('Location', 'Location'),`
- After `text('AnydeskId', 'AnyDesk ID'),`, add:

```js
  text('SerialNumber', 'Serial Number'),
  choice('Status', 'Status', STATUSES),
  date('StatusChangedOn', 'Status Changed On'),
```

- In `CHANGE_COLUMNS`, add `num('DeviceId', 'Device ID'),` as the first entry.
- In `TRACKED_FIELDS`, add `'computerName', 'location', 'serialNumber', 'status',` at the start, and extend the doc comment with: `computerName, serialNumber and status joined when machines got an identity beyond their name: a rename, a first serial and a move to the stash are exactly what the history is for.`
- In `fromListItem`, after the loop and before `return record;`, add:

```js
  // SharePoint's own creation stamp: the earliest this register knew of the
  // machine, which is what "since at least" means on a pre-history owner.
  record.createdOn = row.Created ? new Date(row.Created).getTime() : null;
```

In `src/features/devices/sharepoint/updateDevice.js`:
- `export const EDITABLE_FIELDS = ['owner', 'department', 'deviceType', 'location'];`
- `const COLUMN_FOR = { owner: 'Owner', department: 'Department', deviceType: 'DeviceType', location: 'Location' };`
- Rename the private function to an export and change its signature. Replace `async function logChanges(siteUrl, token, digest, computerName, changes, changedBy) {` with `export async function logChanges(siteUrl, token, digest, device, changes, changedBy) {`. Inside its body, replace `Title: computerName,` with:

```js
          Title: device.computerName ?? '',
          DeviceId: device.id,
```

- Update its two callers: in `updateDevice`, `await logChanges(siteUrl, token, digest, existing, changes, changedBy);`. In `removeOne`, `await logChanges(siteUrl, token, digest, device, [{ … }], changedBy);`.

- [ ] **Step 4: Run the device suite**

Run: `npx vitest run src/features/devices`
Expected: PASS, except possibly `provisionLists.test.js` and `syncPhases.test.js`. They count columns from `DEVICE_COLUMNS` and `CHANGE_COLUMNS` directly, so they should still pass. If one fails, read the assertion: it must be a count derived from the arrays, not a literal.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/lifecycle src/features/devices/sharepoint
git commit -m "Give every machine a status, a location and a serial column, and tie change rows to the machine id

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: The owner-history list

**Files:**
- Create: `src/features/devices/sharepoint/assignmentSchema.js`, `src/features/devices/sharepoint/assignmentSchema.test.js`
- Create: `src/features/devices/sharepoint/readAssignments.js`, `src/features/devices/sharepoint/readAssignments.test.js`
- Modify: `src/features/devices/sharepoint/provisionLists.js`, `src/features/devices/sharepoint/provisionLists.test.js:241-270`, `src/features/devices/sharepoint/syncPhases.test.js` (fake fields)
- Modify: `src/features/devices/sharepoint/readDevices.js`

**Interfaces:**
- Produces:
  - `ASSIGNMENT_LIST_NAME = 'IT Device Assignments'`, `END_REASONS`, `ASSIGNMENT_COLUMNS`
  - A stint is `{ id, deviceId, computerName, owner, location, department, assignedOn, assignedOnApprox, endedOn, endReason, note, recordedBy }`. Times are epoch ms or null.
  - `toAssignmentItem(stint) → object`
  - `closeBody({ endedOn, endReason, note }) → object`
  - `fromAssignmentItem(row) → stint`
  - `readAssignments(siteUrl, token, { deviceId } = {}) → Promise<stint[]>`
  - `readDevice(siteUrl, token, id) → Promise<device|null>`

- [ ] **Step 1: Write the failing tests**

Create `src/features/devices/sharepoint/assignmentSchema.test.js`:

```js
import { describe, it, expect } from 'vitest';
import {
  ASSIGNMENT_LIST_NAME, ASSIGNMENT_COLUMNS, END_REASONS,
  toAssignmentItem, fromAssignmentItem, closeBody,
} from './assignmentSchema.js';

const stint = {
  deviceId: 12, computerName: 'AMIR-HP', owner: 'Amir', location: 'F1', department: 'ENGINEERING',
  assignedOn: Date.UTC(2026, 1, 12, 4), assignedOnApprox: false,
  endedOn: null, endReason: null, note: null, recordedBy: 'it@pmw-group.com',
};

describe('assignment schema', () => {
  it('is its own list', () => expect(ASSIGNMENT_LIST_NAME).toBe('IT Device Assignments'));

  it('names exactly the five ways a stint ends', () => {
    expect(END_REASONS).toEqual(['Reassigned', 'Replaced', 'To stash', 'To repair', 'Retired']);
  });

  it('keeps the time of day on both dates', () => {
    for (const name of ['AssignedOn', 'EndedOn']) {
      expect(ASSIGNMENT_COLUMNS.find((c) => c.StaticName === name).kind).toBe('datetime');
    }
  });

  it('writes an open stint without an end', () => {
    const item = toAssignmentItem(stint);
    expect(item).toMatchObject({
      Title: 'AMIR-HP', DeviceId: 12, Owner: 'Amir', Location: 'F1', Department: 'ENGINEERING',
      AssignedOn: '2026-02-12T04:00:00.000Z', AssignedOnApprox: false, RecordedBy: 'it@pmw-group.com',
    });
    expect(item).not.toHaveProperty('EndedOn');
    expect(item).not.toHaveProperty('EndReason');
  });

  it('round-trips through SharePoint', () => {
    const back = fromAssignmentItem({ Id: 5, ...toAssignmentItem({ ...stint, endedOn: Date.UTC(2026, 9, 5), endReason: 'Reassigned' }) });
    expect(back).toMatchObject({ ...stint, id: 5, endedOn: Date.UTC(2026, 9, 5), endReason: 'Reassigned' });
  });

  it('closes a stint with only the closing columns', () => {
    expect(closeBody({ endedOn: Date.UTC(2026, 9, 5), endReason: 'To stash', note: null }))
      .toEqual({ EndedOn: '2026-10-05T00:00:00.000Z', EndReason: 'To stash' });
  });
});
```

Create `src/features/devices/sharepoint/readAssignments.test.js`:

```js
import { describe, it, expect, afterEach, vi } from 'vitest';
import { readAssignments } from './readAssignments.js';
import { readDevice } from './readDevices.js';

const reply = (body, status = 200) => ({ ok: status < 300, status, json: async () => body });

describe('readAssignments', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('reads every page', async () => {
    const urls = [];
    vi.stubGlobal('fetch', async (url) => {
      urls.push(url);
      return urls.length === 1
        ? reply({ d: { results: [{ Id: 1, DeviceId: 4, Owner: 'Amir' }], __next: 'https://x/next' } })
        : reply({ d: { results: [{ Id: 2, DeviceId: 4, Owner: 'Farid' }] } });
    });
    const stints = await readAssignments('https://x', 't');
    expect(stints.map((s) => s.owner)).toEqual(['Amir', 'Farid']);
  });

  it('asks for one machine when given its id', async () => {
    let asked = '';
    vi.stubGlobal('fetch', async (url) => { asked = url; return reply({ d: { results: [] } }); });
    await readAssignments('https://x', 't', { deviceId: 9 });
    expect(decodeURIComponent(asked)).toContain('$filter=DeviceId eq 9');
  });

  it('reads a list that does not exist yet as no history', async () => {
    vi.stubGlobal('fetch', async () => reply({}, 404));
    await expect(readAssignments('https://x', 't')).resolves.toEqual([]);
  });
});

describe('readDevice', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('reads one row by id', async () => {
    vi.stubGlobal('fetch', async () => reply({ d: { Id: 7, Title: 'PC1', Status: 'Spare' } }));
    const device = await readDevice('https://x', 't', 7);
    expect(device).toMatchObject({ id: 7, computerName: 'PC1', status: 'Spare' });
  });

  it('answers null for a row that has gone', async () => {
    vi.stubGlobal('fetch', async () => reply({}, 404));
    await expect(readDevice('https://x', 't', 7)).resolves.toBeNull();
  });
});
```

In `src/features/devices/sharepoint/provisionLists.test.js`:
- Import `ASSIGNMENT_COLUMNS` from `./assignmentSchema.js`.
- In the `provisionLists progress` describe, change `const TOTAL = DEVICE_COLUMNS.length + CHANGE_COLUMNS.length;` to `const TOTAL = DEVICE_COLUMNS.length + CHANGE_COLUMNS.length + ASSIGNMENT_COLUMNS.length;`.
- Change `[...DEVICE_COLUMNS, ...CHANGE_COLUMNS].map(` in the second test to `[...DEVICE_COLUMNS, ...CHANGE_COLUMNS, ...ASSIGNMENT_COLUMNS].map(`.
- Rename `'reports one tick per column across both lists'` to `'reports one tick per column across all three lists'`.

In `src/features/devices/sharepoint/syncPhases.test.js`:
- Import `ASSIGNMENT_COLUMNS` from `./assignmentSchema.js`.
- Change the `fields` line to `const fields = [...DEVICE_COLUMNS, ...CHANGE_COLUMNS, ...ASSIGNMENT_COLUMNS].map((column) => ({`.

- [ ] **Step 2: Run them to see them fail**

Run: `npx vitest run src/features/devices/sharepoint`
Expected: FAIL. The modules don't exist.

- [ ] **Step 3: Implement**

Create `src/features/devices/sharepoint/assignmentSchema.js`:

```js
/**
 * One row per STINT: somebody had this machine from one date to another.
 *
 * This list is the history; the device row's Owner / Department / Location
 * are the present. A machine therefore never loses its record by changing
 * hands -- "Carmen's old laptop now used by Aisyah" is the next row here.
 *
 * `DeviceId`, not the computer name, ties a stint to its machine, so a rename
 * does not cut the history in two. `Title` carries the name as it was, so the
 * list reads in SharePoint without a join.
 */
export const ASSIGNMENT_LIST_NAME = 'IT Device Assignments';

export const END_REASONS = ['Reassigned', 'Replaced', 'To stash', 'To repair', 'Retired'];

const text = (StaticName, Title) => ({ StaticName, Title, kind: 'text' });
const note = (StaticName, Title) => ({ StaticName, Title, kind: 'note' });
const num = (StaticName, Title) => ({ StaticName, Title, kind: 'number' });
const bool = (StaticName, Title) => ({ StaticName, Title, kind: 'boolean' });
const date = (StaticName, Title) => ({ StaticName, Title, kind: 'datetime' });
const choice = (StaticName, Title, choices) => ({ StaticName, Title, kind: 'choice', choices });

export const ASSIGNMENT_COLUMNS = [
  num('DeviceId', 'Device ID'),
  text('Owner', 'Owner'),
  text('Location', 'Location'),
  text('Department', 'Department'),
  date('AssignedOn', 'Assigned On'),
  bool('AssignedOnApprox', 'Start Is Approximate'),
  date('EndedOn', 'Ended On'),
  choice('EndReason', 'End Reason', END_REASONS),
  note('Note', 'Note'),
  text('RecordedBy', 'Recorded By'),
];

const iso = (ms) => (typeof ms === 'number' && Number.isFinite(ms) ? new Date(ms).toISOString() : null);
const ms = (value) => (value ? new Date(value).getTime() : null);

export function toAssignmentItem(stint) {
  const item = {
    Title: stint.computerName ?? '',
    DeviceId: stint.deviceId,
    Owner: stint.owner ?? '',
    Location: stint.location ?? '',
    Department: stint.department ?? '',
    AssignedOnApprox: Boolean(stint.assignedOnApprox),
    Note: stint.note ?? '',
    RecordedBy: stint.recordedBy ?? '',
  };
  const assignedOn = iso(stint.assignedOn);
  if (assignedOn) item.AssignedOn = assignedOn;
  const endedOn = iso(stint.endedOn);
  if (endedOn) item.EndedOn = endedOn;
  if (stint.endReason) item.EndReason = stint.endReason;
  return item;
}

/** A partial write: only what closing a stint changes. */
export function closeBody({ endedOn, endReason, note: text }) {
  const body = { EndedOn: iso(endedOn), EndReason: endReason };
  if (text) body.Note = text;
  return body;
}

export function fromAssignmentItem(row) {
  const deviceId = row.DeviceId === null || row.DeviceId === undefined || row.DeviceId === ''
    ? null
    : Number(row.DeviceId);
  return {
    id: row.Id ?? row.ID ?? null,
    deviceId,
    computerName: row.Title || null,
    owner: row.Owner || null,
    location: row.Location || null,
    department: row.Department || null,
    assignedOn: ms(row.AssignedOn),
    assignedOnApprox: row.AssignedOnApprox === true,
    endedOn: ms(row.EndedOn),
    endReason: row.EndReason || null,
    note: row.Note || null,
    recordedBy: row.RecordedBy || null,
  };
}
```

Create `src/features/devices/sharepoint/readAssignments.js`:

```js
import { spFetch, listPath } from '../../sharepoint/spClient.js';
import { ASSIGNMENT_LIST_NAME, fromAssignmentItem } from './assignmentSchema.js';

const PAGE_SIZE = 500;

/**
 * Every stint, or one machine's. A list that does not exist yet is the state
 * before the first lifecycle write and reads as no history, not as a failure.
 */
export async function readAssignments(siteUrl, token, { deviceId } = {}) {
  const filter = deviceId === undefined || deviceId === null
    ? ''
    : `&$filter=${encodeURIComponent(`DeviceId eq ${Number(deviceId)}`)}`;
  let url = `${siteUrl}${listPath(ASSIGNMENT_LIST_NAME)}/items?$top=${PAGE_SIZE}${filter}`;
  const rows = [];

  while (url) {
    const response = await spFetch('', url, { token });
    if (response.status === 404) return [];
    if (!response.ok) throw new Error(`Could not read the owner history (${response.status})`);
    const data = await response.json();
    rows.push(...(data.d?.results ?? []));
    url = data.d?.__next ?? null;
  }

  return rows.map(fromAssignmentItem);
}
```

Append to `src/features/devices/sharepoint/readDevices.js`:

```js
/**
 * One machine as SharePoint holds it NOW. A lifecycle action re-reads before
 * it plans, the same rule as a handover: two people moving the same laptop
 * from two screens must not both succeed.
 */
export async function readDevice(siteUrl, token, id) {
  const response = await spFetch(siteUrl, `${listPath(DEVICE_LIST_NAME)}/items(${Number(id)})`, { token });
  if (response.status === 404) return null;
  if (!response.ok) throw new Error(`Could not read that machine (${response.status})`);
  const data = await response.json();
  return refixStored(fromListItem(data.d ?? data));
}
```

In `src/features/devices/sharepoint/provisionLists.js`:
- Import `{ ASSIGNMENT_LIST_NAME, ASSIGNMENT_COLUMNS } from './assignmentSchema.js';`
- Append to the `lists` array:

```js
      {
        title: ASSIGNMENT_LIST_NAME,
        description: 'Who had each machine, from when to when',
        columns: ASSIGNMENT_COLUMNS,
      },
```

- Change "across both lists" in the doc comment to "across all three lists".

- [ ] **Step 4: Run the device suite**

Run: `npx vitest run src/features/devices`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/sharepoint
git commit -m "Add the IT Device Assignments list: one row per stint of somebody having a machine

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: Stints and the owner timeline

**Files:**
- Create: `src/features/devices/lifecycle/stints.js`, `src/features/devices/lifecycle/stints.test.js`

**Interfaces:**
- Consumes: `inFleet`, `statusOf`, `SPARE`, `RETIRED` (Task 3); the stint shape (Task 4).
- Produces:
  - `stintsOf(deviceId, stints) → stint[]` (newest first)
  - `openStintOf(deviceId, stints) → stint|null`
  - `legacyStint(device) → stint|null` (`id: null`, `assignedOnApprox: true`)
  - `currentStint(device, stints) → stint|null` (the stint an action would end)
  - `timelineOf(device, stints) → Array<{ kind: 'stint', …stint, current } | { kind: 'gap', from, to, label }>` (newest first)

- [ ] **Step 1: Write the failing test**

Create `src/features/devices/lifecycle/stints.test.js`:

```js
import { describe, it, expect } from 'vitest';
import {
  stintsOf, openStintOf, legacyStint, currentStint, timelineOf,
} from './stints.js';

const day = (d) => Date.UTC(2026, 0, d);
const machine = (over) => ({
  id: 4, computerName: 'AMIR-HP', owner: 'Amir', location: 'F1', department: 'ENGINEERING',
  status: null, createdOn: day(1), importedOn: day(2), ...over,
});
const stint = (over) => ({
  id: 1, deviceId: 4, owner: 'Farid', assignedOn: day(3), endedOn: day(10), endReason: 'Replaced', ...over,
});

describe('stintsOf / openStintOf', () => {
  it('keeps only this machine, newest first', () => {
    const list = [stint({ id: 1 }), stint({ id: 2, assignedOn: day(12), endedOn: null }), stint({ id: 3, deviceId: 9 })];
    expect(stintsOf(4, list).map((s) => s.id)).toEqual([2, 1]);
    expect(openStintOf(4, list).id).toBe(2);
    expect(openStintOf(9, [stint({ deviceId: 9 })])).toBeNull();
  });
});

describe('legacyStint', () => {
  it('is what the register implies about a machine with no history: since at least its creation', () => {
    expect(legacyStint(machine())).toMatchObject({
      id: null, deviceId: 4, owner: 'Amir', location: 'F1', department: 'ENGINEERING',
      assignedOn: day(1), assignedOnApprox: true, endedOn: null,
    });
  });

  it('falls back to the import date when there is no creation stamp', () => {
    expect(legacyStint(machine({ createdOn: null })).assignedOn).toBe(day(2));
  });

  it('is nothing for a machine with no owner', () => {
    expect(legacyStint(machine({ owner: null }))).toBeNull();
  });
});

describe('currentStint', () => {
  it('is the open stint when there is one', () => {
    expect(currentStint(machine(), [stint({ endedOn: null })]).id).toBe(1);
  });

  it('is the legacy stint for an in-use machine with no history at all', () => {
    expect(currentStint(machine(), []).id).toBeNull();
  });

  it('is nothing for a machine whose history has ended', () => {
    expect(currentStint(machine(), [stint()])).toBeNull();
  });

  it('is nothing for a spare machine', () => {
    expect(currentStint(machine({ status: 'Spare' }), [])).toBeNull();
  });
});

describe('timelineOf', () => {
  it('shows the time in the stash between two people', () => {
    const list = [
      stint({ id: 2, owner: 'Amir', assignedOn: day(20), endedOn: null, endReason: null }),
      stint({ id: 1, owner: 'Farid', assignedOn: day(3), endedOn: day(10), endReason: 'To stash' }),
    ];
    const entries = timelineOf(machine(), list);
    expect(entries.map((e) => e.kind)).toEqual(['stint', 'gap', 'stint']);
    expect(entries[0].current).toBe(true);
    expect(entries[1]).toEqual({ kind: 'gap', from: day(10), to: day(20), label: 'In IT Stash' });
  });

  it('ends with the stash or the graveyard for a machine nobody has now', () => {
    const entries = timelineOf(machine({ status: 'Retired', owner: null }), [stint({ endReason: 'Retired' })]);
    expect(entries[0]).toEqual({ kind: 'gap', from: day(10), to: null, label: 'In the Graveyard' });
  });

  it('shows the implied owner for a machine with no history yet', () => {
    const entries = timelineOf(machine(), []);
    expect(entries).toHaveLength(1);
    expect(entries[0]).toMatchObject({ kind: 'stint', owner: 'Amir', assignedOnApprox: true, current: true });
  });
});
```

- [ ] **Step 2: Run it to see it fail**

Run: `npx vitest run src/features/devices/lifecycle/stints.test.js`
Expected: FAIL. The module doesn't exist.

- [ ] **Step 3: Implement**

Create `src/features/devices/lifecycle/stints.js`:

```js
import { inFleet, statusOf, SPARE, RETIRED } from './status.js';

const newestFirst = (a, b) => (b.assignedOn ?? 0) - (a.assignedOn ?? 0);

export function stintsOf(deviceId, stints = []) {
  return stints.filter((stint) => stint.deviceId === deviceId).sort(newestFirst);
}

export function openStintOf(deviceId, stints = []) {
  return stintsOf(deviceId, stints).find((stint) => stint.endedOn === null || stint.endedOn === undefined) ?? null;
}

/**
 * What the register already says about a machine nobody has written history
 * for: its current owner, since at least the day the row appeared. Never
 * written until something ends it -- nothing is written for a machine nobody
 * touches.
 */
export function legacyStint(device) {
  if (!device?.owner) return null;
  return {
    id: null,
    deviceId: device.id,
    computerName: device.computerName ?? null,
    owner: device.owner,
    location: device.location ?? null,
    department: device.department ?? null,
    assignedOn: device.createdOn ?? device.importedOn ?? device.scannedOn ?? null,
    assignedOnApprox: true,
    endedOn: null,
    endReason: null,
    note: null,
    recordedBy: null,
  };
}

/** The stint a lifecycle action would end, or null when nobody has the machine. */
export function currentStint(device, stints = []) {
  const open = openStintOf(device.id, stints);
  if (open) return open;
  if (!inFleet(device)) return null;
  // History exists and none of it is open: the register's owner is not backed
  // by a stint, and inventing one now would put a guess after real records.
  if (stintsOf(device.id, stints).length) return null;
  return legacyStint(device);
}

const gapLabel = (reason) => {
  if (reason === 'To stash') return 'In IT Stash';
  if (reason === 'Retired') return 'In the Graveyard';
  return 'Nobody recorded';
};

const nowLabel = (device) => {
  if (statusOf(device) === SPARE) return 'In IT Stash';
  if (statusOf(device) === RETIRED) return 'In the Graveyard';
  return 'Nobody recorded';
};

/** Owner history newest first, with the spells nobody had it shown as gaps. */
export function timelineOf(device, stints = []) {
  const own = stintsOf(device.id, stints);
  const list = own.length ? own : [legacyStint(device)].filter(Boolean);
  const entries = [];

  const latest = list[0];
  if (latest && latest.endedOn !== null && latest.endedOn !== undefined) {
    entries.push({ kind: 'gap', from: latest.endedOn, to: null, label: nowLabel(device) });
  }

  list.forEach((stint, index) => {
    entries.push({ kind: 'stint', ...stint, current: stint.endedOn === null || stint.endedOn === undefined });
    const older = list[index + 1];
    if (older?.endedOn != null && stint.assignedOn != null && stint.assignedOn > older.endedOn) {
      entries.push({ kind: 'gap', from: older.endedOn, to: stint.assignedOn, label: gapLabel(older.endReason) });
    }
  });

  return entries;
}
```

- [ ] **Step 4: Run it to see it pass**

Run: `npx vitest run src/features/devices/lifecycle/stints.test.js`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/lifecycle/stints.js src/features/devices/lifecycle/stints.test.js
git commit -m "Work out a machine's current holder and its owner timeline, including spells in the stash

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: What each lifecycle action does

**Files:**
- Create: `src/features/devices/lifecycle/planLifecycle.js`, `src/features/devices/lifecycle/planLifecycle.test.js`

**Interfaces:**
- Consumes: `statusOf`, the status constants (Task 3); `currentStint` (Task 5); `cleanLocation` (Task 2).
- Produces:
  - `ACTIONS = { CHANGE_OWNER: 'changeOwner', TO_REPAIR: 'toRepair', BACK_IN_USE: 'backInUse', TO_STASH: 'toStash', RETIRE: 'retire', ASSIGN: 'assign' }`
  - `actionsFor(device) → string[]`
  - `planLifecycle({ device, stints, action, input, expectedStatus, now, recordedBy })` returns `{ refusal }` or `{ fields, changes, close, open }`, where:
    - `fields`: camelCase device fields to MERGE (`owner`, `location`, `department`, `ownerSource`, `status`, `statusChangedOn`, `manualFields`)
    - `changes`: `[{ fieldName, oldValue, newValue, changeType }]`
    - `close`: `{ stint, endedOn, endReason, note } | null`
    - `open`: stint without `id` `| null`
  - `input`: `{ owner?, location?, department?, on?: ms, note?, endReason? }`. `endReason` overrides the reason for `TO_STASH` / `RETIRE`; the import uses it for `Replaced`.

- [ ] **Step 1: Write the failing test**

Create `src/features/devices/lifecycle/planLifecycle.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { ACTIONS, actionsFor, planLifecycle } from './planLifecycle.js';

const NOW = Date.UTC(2026, 9, 5, 4);
const machine = (over) => ({
  id: 4, computerName: 'AMIR-HP', owner: 'Amir', location: 'F1', department: 'ENGINEERING',
  status: null, manualFields: [], createdOn: Date.UTC(2026, 7, 21), ...over,
});
const open = { id: 30, deviceId: 4, owner: 'Amir', assignedOn: Date.UTC(2026, 1, 12), endedOn: null };
const plan = (over) => planLifecycle({ device: machine(), stints: [open], now: NOW, recordedBy: 'it', ...over });

describe('actionsFor', () => {
  it('offers what each status allows', () => {
    expect(actionsFor(machine())).toEqual(['changeOwner', 'toRepair', 'toStash', 'retire']);
    expect(actionsFor(machine({ status: 'In repair' }))).toEqual(['changeOwner', 'backInUse', 'toStash', 'retire']);
    expect(actionsFor(machine({ status: 'Spare' }))).toEqual(['assign', 'retire']);
    expect(actionsFor(machine({ status: 'Retired' }))).toEqual(['assign']);
  });
});

describe('planLifecycle — change owner', () => {
  const result = plan({
    action: ACTIONS.CHANGE_OWNER,
    input: { owner: 'Aisyah', location: 'f1', department: 'ENGINEERING', note: 'Amir got a workstation' },
  });

  it('ends the current stint as Reassigned, with the note', () => {
    expect(result.close).toEqual({ stint: open, endedOn: NOW, endReason: 'Reassigned', note: 'Amir got a workstation' });
  });

  it('opens one for the new holder', () => {
    expect(result.open).toMatchObject({
      deviceId: 4, computerName: 'AMIR-HP', owner: 'Aisyah', location: 'F1', department: 'ENGINEERING',
      assignedOn: NOW, assignedOnApprox: false, endedOn: null, note: null, recordedBy: 'it',
    });
  });

  it('marks the holder as set by hand so an import leaves it alone', () => {
    expect(result.fields).toMatchObject({ owner: 'Aisyah', ownerSource: 'Manual' });
    expect(result.fields.manualFields.sort()).toEqual(['department', 'location', 'owner']);
    expect(result.changes).toEqual([
      { fieldName: 'owner', oldValue: 'Amir', newValue: 'Aisyah', changeType: 'Updated' },
    ]);
  });

  it('refuses a change to the same holder', () => {
    expect(plan({ action: ACTIONS.CHANGE_OWNER, input: { owner: 'amir', location: 'F1', department: 'engineering' } }).refusal)
      .toMatch(/already has/);
  });

  it('refuses without a name', () => {
    expect(plan({ action: ACTIONS.CHANGE_OWNER, input: { owner: '  ' } }).refusal).toMatch(/Type who/);
  });

  it('treats a move to another location as a new stint', () => {
    const moved = plan({ action: ACTIONS.CHANGE_OWNER, input: { owner: 'Amir', location: 'F3', department: 'ENGINEERING' } });
    expect(moved.refusal).toBeUndefined();
    expect(moved.changes).toEqual([{ fieldName: 'location', oldValue: 'F1', newValue: 'F3', changeType: 'Updated' }]);
  });
});

describe('planLifecycle — repair', () => {
  it('keeps the stint open: it is still their laptop', () => {
    const result = plan({ action: ACTIONS.TO_REPAIR });
    expect(result.close).toBeNull();
    expect(result.open).toBeNull();
    expect(result.fields).toMatchObject({ status: 'In repair', statusChangedOn: NOW });
  });

  it('comes back into use', () => {
    expect(plan({ device: machine({ status: 'In repair' }), action: ACTIONS.BACK_IN_USE }).fields.status).toBe('In use');
  });
});

describe('planLifecycle — stash and retire', () => {
  it('ends the stint and clears the holder going to the stash', () => {
    const result = plan({ device: machine({ manualFields: ['owner', 'deviceType'] }), action: ACTIONS.TO_STASH });
    expect(result.close.endReason).toBe('To stash');
    expect(result.fields).toMatchObject({ status: 'Spare', owner: null, department: null });
    expect(result.fields.manualFields).toEqual(['deviceType']);
  });

  it('records a replacement as Replaced, not To stash', () => {
    expect(plan({ action: ACTIONS.TO_STASH, input: { endReason: 'Replaced' } }).close.endReason).toBe('Replaced');
  });

  it('retires, ending the legacy stint of a machine with no history', () => {
    const result = plan({ stints: [], action: ACTIONS.RETIRE });
    expect(result.close.stint).toMatchObject({ id: null, owner: 'Amir', assignedOnApprox: true });
    expect(result.close.endReason).toBe('Retired');
    expect(result.fields.status).toBe('Retired');
  });

  it('brings a retired machine back for somebody', () => {
    const result = plan({
      device: machine({ status: 'Retired', owner: null, department: null }),
      stints: [],
      action: ACTIONS.ASSIGN,
      input: { owner: 'Hafizah', location: 'PML', department: 'SALES', note: 'loan' },
    });
    expect(result.close).toBeNull();
    expect(result.open).toMatchObject({ owner: 'Hafizah', location: 'PML', note: 'loan' });
    expect(result.fields.status).toBe('In use');
  });
});

describe('planLifecycle — refusals', () => {
  it('refuses when somebody else moved the machine first', () => {
    expect(plan({ device: machine({ status: 'Spare' }), action: ACTIONS.ASSIGN, input: { owner: 'X' }, expectedStatus: 'In use' }).refusal)
      .toMatch(/just moved to IT Stash by someone else/);
  });

  it('refuses an action the status does not allow', () => {
    expect(plan({ action: ACTIONS.ASSIGN, input: { owner: 'X' } }).refusal).toMatch(/cannot/);
  });

  it('uses the effective date given', () => {
    const on = Date.UTC(2026, 8, 30, 4);
    expect(plan({ action: ACTIONS.TO_STASH, input: { on } }).close.endedOn).toBe(on);
  });
});
```

- [ ] **Step 2: Run it to see it fail**

Run: `npx vitest run src/features/devices/lifecycle/planLifecycle.test.js`
Expected: FAIL.

- [ ] **Step 3: Implement**

Create `src/features/devices/lifecycle/planLifecycle.js`:

```js
import {
  statusOf, IN_USE, IN_REPAIR, SPARE, RETIRED,
} from './status.js';
import { currentStint } from './stints.js';
import { cleanLocation } from '../map/locations.js';

/**
 * What one action on one machine writes. Pure: the caller re-reads the
 * machine and its stints immediately before calling, and performs the result
 * in order -- device row, then stints, then change log.
 */
export const ACTIONS = {
  CHANGE_OWNER: 'changeOwner',
  TO_REPAIR: 'toRepair',
  BACK_IN_USE: 'backInUse',
  TO_STASH: 'toStash',
  RETIRE: 'retire',
  ASSIGN: 'assign',
};

const ALLOWED = {
  [IN_USE]: [ACTIONS.CHANGE_OWNER, ACTIONS.TO_REPAIR, ACTIONS.TO_STASH, ACTIONS.RETIRE],
  [IN_REPAIR]: [ACTIONS.CHANGE_OWNER, ACTIONS.BACK_IN_USE, ACTIONS.TO_STASH, ACTIONS.RETIRE],
  [SPARE]: [ACTIONS.ASSIGN, ACTIONS.RETIRE],
  [RETIRED]: [ACTIONS.ASSIGN],
};

export function actionsFor(device) {
  return ALLOWED[statusOf(device)];
}

const WHERE = { [SPARE]: 'IT Stash', [RETIRED]: 'the Graveyard', [IN_REPAIR]: 'repair', [IN_USE]: 'use' };
const text = (value) => (value === null || value === undefined ? '' : String(value).trim());
const same = (a, b) => text(a).toUpperCase() === text(b).toUpperCase();

export function planLifecycle({
  device, stints = [], action, input = {}, expectedStatus, now = Date.now(), recordedBy = '',
}) {
  const status = statusOf(device);

  if (expectedStatus && expectedStatus !== status) {
    return { refusal: `This machine was just moved to ${WHERE[status]} by someone else. Refresh to see it as it is now.` };
  }
  if (!ALLOWED[status].includes(action)) {
    return { refusal: `A machine that is ${status.toLowerCase()} cannot do that.` };
  }

  const on = typeof input.on === 'number' ? input.on : now;
  const note = text(input.note) || null;
  const stint = currentStint(device, stints);
  const manual = new Set(device.manualFields ?? []);
  const fields = {};
  const changes = [];
  let close = null;
  let open = null;

  const set = (key, value) => {
    const before = text(device[key]);
    const after = text(value);
    if (before === after) return;
    fields[key] = value;
    let changeType = 'Updated';
    if (!before) changeType = 'Added';
    else if (!after) changeType = 'Removed';
    changes.push({ fieldName: key, oldValue: before, newValue: after, changeType });
  };
  const setStatus = (next) => {
    set('status', next);
    fields.statusChangedOn = on;
  };
  const holder = () => {
    const owner = text(input.owner);
    if (!owner) return null;
    return { owner, location: cleanLocation(input.location), department: text(input.department) || null };
  };
  const take = (who) => {
    set('owner', who.owner);
    set('location', who.location);
    set('department', who.department);
    fields.ownerSource = 'Manual';
    ['owner', 'location', 'department'].forEach((key) => manual.add(key));
  };
  const release = () => {
    set('owner', null);
    set('department', null);
    // Blank and handed back to the scan: a later scan naming somebody is how a
    // spare machine comes back into use without anyone pressing anything.
    manual.delete('owner');
    manual.delete('department');
  };
  const end = (reason) => {
    if (stint) close = { stint, endedOn: on, endReason: reason, note };
  };
  const start = (who) => {
    open = {
      deviceId: device.id,
      computerName: device.computerName ?? null,
      ...who,
      assignedOn: on,
      assignedOnApprox: false,
      endedOn: null,
      endReason: null,
      // A note explains why a stint ENDED when one ends; otherwise it belongs
      // to the one starting.
      note: close ? null : note,
      recordedBy,
    };
  };

  switch (action) {
    case ACTIONS.CHANGE_OWNER: {
      const who = holder();
      if (!who) return { refusal: 'Type who the machine is going to.' };
      if (same(who.owner, device.owner) && same(who.location, device.location) && same(who.department, device.department)) {
        return { refusal: `${who.owner} already has this machine.` };
      }
      end('Reassigned');
      take(who);
      start(who);
      break;
    }
    case ACTIONS.TO_REPAIR:
      setStatus(IN_REPAIR);
      break;
    case ACTIONS.BACK_IN_USE:
      setStatus(IN_USE);
      break;
    case ACTIONS.TO_STASH:
      end(input.endReason ?? 'To stash');
      release();
      setStatus(SPARE);
      break;
    case ACTIONS.RETIRE:
      end(input.endReason ?? 'Retired');
      release();
      setStatus(RETIRED);
      break;
    case ACTIONS.ASSIGN: {
      const who = holder();
      if (!who) return { refusal: 'Type who the machine is going to.' };
      take(who);
      setStatus(IN_USE);
      start(who);
      break;
    }
    default:
      return { refusal: 'That is not something a machine can do.' };
  }

  fields.manualFields = [...manual];
  return { fields, changes, close, open };
}
```

- [ ] **Step 4: Run it to see it pass**

Run: `npx vitest run src/features/devices/lifecycle`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/lifecycle/planLifecycle.js src/features/devices/lifecycle/planLifecycle.test.js
git commit -m "Decide what changing owner, repair, stash, retire and bring-back each write, and refuse a stale screen

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 7: Performing a lifecycle action against SharePoint

**Files:**
- Create: `src/features/devices/sharepoint/writeLifecycle.js`, `src/features/devices/sharepoint/writeLifecycle.test.js`

**Interfaces:**
- Consumes:
  - `planLifecycle` (Task 6)
  - `readDevice`, `readAssignments`, `toAssignmentItem`, `closeBody`, `ASSIGNMENT_LIST_NAME` (Task 4)
  - `logChanges` (Task 3)
- Produces:
  - `lifecycleItem(fields) → SharePoint MERGE body`
  - `stintWrites({ close, open }) → Array<{ id: number|null, body }>`
  - `writeStints({ siteUrl, token, digest, writes })`. On failure it throws an Error with `.remaining`, the writes not yet done.
  - `performLifecycle({ siteUrl, token, deviceId, action, input, expectedStatus, recordedBy, now }) → { plan, pendingStints }`. It throws on a refusal or a device-row failure.

- [ ] **Step 1: Write the failing test**

Create `src/features/devices/sharepoint/writeLifecycle.test.js`:

```js
import { describe, it, expect, afterEach, vi } from 'vitest';
import { lifecycleItem, stintWrites, performLifecycle } from './writeLifecycle.js';

const SITE = 'https://contoso.sharepoint.com/sites/it';
const NOW = Date.UTC(2026, 9, 5, 4);

describe('lifecycleItem', () => {
  it('maps only the fields the action changed', () => {
    expect(lifecycleItem({
      owner: null, department: null, status: 'Spare', statusChangedOn: NOW, manualFields: ['deviceType'],
    })).toEqual({
      Owner: '', Department: '', Status: 'Spare',
      StatusChangedOn: '2026-10-05T04:00:00.000Z', ManualFields: 'deviceType',
    });
  });
});

describe('stintWrites', () => {
  it('merges the end onto a stint that exists', () => {
    const writes = stintWrites({ close: { stint: { id: 30 }, endedOn: NOW, endReason: 'To stash', note: null }, open: null });
    expect(writes).toEqual([{ id: 30, body: { EndedOn: '2026-10-05T04:00:00.000Z', EndReason: 'To stash' } }]);
  });

  it('writes a legacy stint whole, already ended', () => {
    const [write] = stintWrites({
      close: { stint: { id: null, deviceId: 4, owner: 'Amir', assignedOn: NOW - 1000, assignedOnApprox: true }, endedOn: NOW, endReason: 'Retired', note: 'old' },
      open: null,
    });
    expect(write.id).toBeNull();
    expect(write.body).toMatchObject({ DeviceId: 4, Owner: 'Amir', AssignedOnApprox: true, EndReason: 'Retired', Note: 'old' });
  });
});

function fakeSharePoint({ device, stints = [], failStints = false } = {}) {
  const calls = [];
  const reply = (body, status = 200) => ({
    ok: status < 300, status, json: async () => body, text: async () => 'boom', headers: { get: () => null },
  });
  return {
    calls,
    fetch: async (url, init = {}) => {
      const method = init.method ?? 'GET';
      calls.push({ url: decodeURIComponent(url), method, headers: init.headers ?? {}, body: init.body ? JSON.parse(init.body) : undefined });
      if (url.endsWith('/_api/contextinfo')) return reply({ d: { GetContextWebInformation: { FormDigestValue: 'D' } } });
      if (method === 'GET' && url.includes('IT%20Device%20List')) return reply({ d: device });
      if (method === 'GET' && url.includes('IT%20Device%20Assignments')) return reply({ d: { results: stints } });
      if (failStints && url.includes('IT%20Device%20Assignments')) return reply({}, 500);
      return reply({ Id: 99 }, 201);
    },
  };
}

const posts = (calls) => calls.filter((c) => c.method === 'POST' && !c.url.endsWith('/contextinfo'));

describe('performLifecycle', () => {
  afterEach(() => vi.unstubAllGlobals());

  const device = { Id: 4, Title: 'AMIR-HP', Owner: 'Amir', Location: 'F1', Department: 'ENGINEERING', Created: '2026-08-21T00:00:00Z' };

  it('writes the device row, then the stints, then the change log', async () => {
    const sp = fakeSharePoint({ device, stints: [{ Id: 30, DeviceId: 4, Owner: 'Amir', AssignedOn: '2026-02-12T00:00:00Z' }] });
    vi.stubGlobal('fetch', sp.fetch);

    await performLifecycle({
      siteUrl: SITE, token: 't', deviceId: 4, action: 'changeOwner',
      input: { owner: 'Aisyah', location: 'F1', department: 'ENGINEERING' }, expectedStatus: 'In use', recordedBy: 'it', now: NOW,
    });

    const order = posts(sp.calls).map((c) => {
      if (c.url.includes('IT Device List')) return 'device';
      if (c.url.includes('IT Device Assignments')) return 'stint';
      return 'change';
    });
    expect(order).toEqual(['device', 'stint', 'stint', 'change']);
    expect(posts(sp.calls)[1].url).toContain('items(30)');
    expect(posts(sp.calls)[3].body).toMatchObject({ DeviceId: 4, FieldName: 'owner', NewValue: 'Aisyah' });
  });

  it('refuses and writes nothing when the screen was stale', async () => {
    const sp = fakeSharePoint({ device: { ...device, Status: 'Spare', Owner: '' } });
    vi.stubGlobal('fetch', sp.fetch);

    await expect(performLifecycle({
      siteUrl: SITE, token: 't', deviceId: 4, action: 'toStash', expectedStatus: 'In use', now: NOW,
    })).rejects.toThrow(/someone else/);
    expect(posts(sp.calls)).toHaveLength(0);
  });

  it('reports stints it could not write instead of failing the whole action', async () => {
    const sp = fakeSharePoint({ device, failStints: true });
    vi.stubGlobal('fetch', sp.fetch);

    const result = await performLifecycle({
      siteUrl: SITE, token: 't', deviceId: 4, action: 'retire', expectedStatus: 'In use', now: NOW,
    });
    expect(result.pendingStints).toHaveLength(1);
  });
});
```

- [ ] **Step 2: Run it to see it fail**

Run: `npx vitest run src/features/devices/sharepoint/writeLifecycle.test.js`
Expected: FAIL.

- [ ] **Step 3: Implement**

Create `src/features/devices/sharepoint/writeLifecycle.js`:

```js
import {
  spFetch, listPath, ITEM_ACCEPT, getFormDigest,
} from '../../sharepoint/spClient.js';
import { withRetry } from '../../sharepoint/writePool.js';
import { DEVICE_LIST_NAME } from './deviceSchema.js';
import { ASSIGNMENT_LIST_NAME, toAssignmentItem, closeBody } from './assignmentSchema.js';
import { readDevice } from './readDevices.js';
import { readAssignments } from './readAssignments.js';
import { logChanges } from './updateDevice.js';
import { planLifecycle } from '../lifecycle/planLifecycle.js';

const MERGE = { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' };

const TEXT_COLUMN = {
  owner: 'Owner', location: 'Location', department: 'Department', ownerSource: 'OwnerSource', status: 'Status',
};

/** A PARTIAL write: only the columns the action changed. Never toListItem. */
export function lifecycleItem(fields) {
  const body = {};
  for (const [key, column] of Object.entries(TEXT_COLUMN)) {
    if (key in fields) body[column] = fields[key] === null || fields[key] === undefined ? '' : String(fields[key]);
  }
  if ('statusChangedOn' in fields) body.StatusChangedOn = new Date(fields.statusChangedOn).toISOString();
  if ('manualFields' in fields) body.ManualFields = fields.manualFields.join('\n');
  return body;
}

/** The owner-history writes a plan implies, in order: end the old, start the new. */
export function stintWrites({ close, open }) {
  const writes = [];
  if (close) {
    writes.push(close.stint.id !== null && close.stint.id !== undefined
      ? { id: close.stint.id, body: closeBody(close) }
      : {
        id: null,
        body: toAssignmentItem({
          ...close.stint, endedOn: close.endedOn, endReason: close.endReason, note: close.note ?? close.stint.note,
        }),
      });
  }
  if (open) writes.push({ id: null, body: toAssignmentItem(open) });
  return writes;
}

export async function writeStints({ siteUrl, token, digest, writes }) {
  for (let index = 0; index < writes.length; index += 1) {
    const write = writes[index];
    const path = write.id === null
      ? `${listPath(ASSIGNMENT_LIST_NAME)}/items`
      : `${listPath(ASSIGNMENT_LIST_NAME)}/items(${write.id})`;
    const response = await withRetry(() => spFetch(siteUrl, path, {
      token, digest, method: 'POST', accept: ITEM_ACCEPT, body: write.body,
      ...(write.id === null ? {} : { headers: MERGE }),
    }));
    if (!response.ok) {
      const failure = new Error(`Could not record the owner history (${response.status})`);
      // Only what has not landed is retried: re-posting an ended legacy stint
      // that already went in would put the same person in the history twice.
      failure.remaining = writes.slice(index);
      throw failure;
    }
  }
}

/**
 * Re-read, plan, write. The device row goes first and is the only write that
 * fails the action: a machine whose status moved but whose history did not is
 * recoverable from `pendingStints`; the reverse would be history describing a
 * move that never happened.
 */
export async function performLifecycle({
  siteUrl, token, deviceId, action, input, expectedStatus, recordedBy = '', now = Date.now(),
}) {
  const digest = await getFormDigest(siteUrl, token);
  const device = await readDevice(siteUrl, token, deviceId);
  if (!device) throw new Error('That machine is no longer in the register.');
  const stints = await readAssignments(siteUrl, token, { deviceId });

  const plan = planLifecycle({ device, stints, action, input, expectedStatus, now, recordedBy });
  if (plan.refusal) throw new Error(plan.refusal);

  const response = await withRetry(() => spFetch(siteUrl, `${listPath(DEVICE_LIST_NAME)}/items(${device.id})`, {
    token, digest, method: 'POST', accept: ITEM_ACCEPT, body: lifecycleItem(plan.fields), headers: MERGE,
  }));
  if (!response.ok) throw new Error(`Could not save the change (${response.status}): ${await response.text()}`);

  let pendingStints = [];
  try {
    await writeStints({ siteUrl, token, digest, writes: stintWrites(plan) });
  } catch (failure) {
    pendingStints = failure.remaining ?? stintWrites(plan);
  }

  await logChanges(siteUrl, token, digest, device, plan.changes, recordedBy);
  return { plan, pendingStints };
}

/** Retry only the history writes an action left behind. */
export async function retryStints({ siteUrl, token, writes }) {
  const digest = await getFormDigest(siteUrl, token);
  await writeStints({ siteUrl, token, digest, writes });
}
```

- [ ] **Step 4: Run it to see it pass**

Run: `npx vitest run src/features/devices/sharepoint/writeLifecycle.test.js`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/sharepoint/writeLifecycle.js src/features/devices/sharepoint/writeLifecycle.test.js
git commit -m "Perform a lifecycle action: re-read, then device row, owner history and change log in that order

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 8: Matching an import, and spotting replacements

**Files:**
- Create: `src/features/devices/lifecycle/matchIncoming.js`, `src/features/devices/lifecycle/matchIncoming.test.js`
- Create: `src/features/devices/lifecycle/replacements.js`, `src/features/devices/lifecycle/replacements.test.js`

**Interfaces:**
- Consumes: `normaliseSerial` (Task 1); `inFleet`, `statusOf` (Task 3).
- Produces:
  - `MATCH = { NEW, SAME, RENAMED, NAME_REUSED, DUPLICATE_SERIAL }` (string values `'new' | 'same' | 'renamed' | 'nameReused' | 'duplicateSerial'`)
  - `matchIncoming(incoming, existing) → Array<{ device, existing: row|null, kind, note: string|null }>`
  - `noticeFor(match) → string|null`
  - `ANSWERS = { STASH: 'stash', RETIRE: 'retire', KEEP: 'keep' }`
  - `replacementsFor(matches, existing) → Array<{ key, sourceFileName, incomingName, owner, old: { id, computerName, deviceType, status, since } }>`
  - `unanswered(prompts, answers, excluded = new Set()) → number`

- [ ] **Step 1: Write the failing tests**

Create `src/features/devices/lifecycle/matchIncoming.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { MATCH, matchIncoming, noticeFor } from './matchIncoming.js';

const scan = (over) => ({ computerName: 'AMIR-HP', serialNumber: '5CG8241KQZ', owner: 'Amir', sourceFileName: 'a.txt', ...over });
const row = (over) => ({ id: 4, computerName: 'AMIR-HP', serialNumber: '5CG8241KQZ', owner: 'Amir', ...over });

describe('matchIncoming', () => {
  it('finds a machine by its serial', () => {
    const [m] = matchIncoming([scan()], [row()]);
    expect(m.kind).toBe(MATCH.SAME);
    expect(m.existing.id).toBe(4);
  });

  it('sees a rename: same serial, new name', () => {
    const [m] = matchIncoming([scan({ computerName: 'AISYAH-HP' })], [row()]);
    expect(m.kind).toBe(MATCH.RENAMED);
    expect(m.note).toMatch(/Renamed from AMIR-HP/);
  });

  it('matches serials whatever their spacing or case', () => {
    expect(matchIncoming([scan({ serialNumber: '5cg8241kqz' })], [row({ serialNumber: '5CG 8241 KQZ' })])[0].kind).toBe(MATCH.SAME);
  });

  it('falls back to the name for a report with no serial', () => {
    expect(matchIncoming([scan({ serialNumber: null })], [row()])[0].kind).toBe(MATCH.SAME);
  });

  it('writes the serial onto a row that had none', () => {
    const [m] = matchIncoming([scan()], [row({ serialNumber: null })]);
    expect(m.kind).toBe(MATCH.SAME);
    expect(m.existing.id).toBe(4);
  });

  it('treats a reused name with a different serial as a new machine', () => {
    const [m] = matchIncoming([scan({ serialNumber: 'NEWSERIAL1' })], [row()]);
    expect(m.kind).toBe(MATCH.NAME_REUSED);
    expect(m.existing).toBeNull();
    expect(m.note).toMatch(/different machine \(serial 5CG8241KQZ\)/);
  });

  it('flags two reports in one drop claiming one serial', () => {
    const matches = matchIncoming([scan(), scan({ computerName: 'OTHER', sourceFileName: 'b.txt' })], []);
    expect(matches.map((m) => m.kind)).toEqual([MATCH.NEW, MATCH.DUPLICATE_SERIAL]);
  });

  it('is new when nothing matches', () => {
    expect(matchIncoming([scan()], [])[0].kind).toBe(MATCH.NEW);
  });
});

describe('noticeFor', () => {
  it('says a retired machine scanned again is being brought back', () => {
    const [m] = matchIncoming([scan()], [row({ status: 'Retired' })]);
    expect(noticeFor(m)).toMatch(/Brought back from the Graveyard/);
  });

  it('says nothing for an ordinary re-scan', () => {
    expect(noticeFor(matchIncoming([scan()], [row()])[0])).toBeNull();
  });
});
```

Create `src/features/devices/lifecycle/replacements.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { matchIncoming } from './matchIncoming.js';
import { replacementsFor, unanswered } from './replacements.js';

const scan = (over) => ({ computerName: 'CARMEN-HP2', serialNumber: 'NEW1', owner: 'Carmen', sourceFileName: 'c.txt', ...over });
const row = (over) => ({ id: 4, computerName: 'CARMEN-HP', serialNumber: 'OLD1', owner: 'Carmen', deviceType: 'Laptop', ...over });
const prompts = (incoming, existing) => replacementsFor(matchIncoming(incoming, existing), existing);

describe('replacementsFor', () => {
  it('asks when a person turns up on a new machine', () => {
    const [p] = prompts([scan()], [row()]);
    expect(p).toMatchObject({ key: 'c.txt|4', owner: 'Carmen', incomingName: 'CARMEN-HP2', old: { id: 4, computerName: 'CARMEN-HP' } });
  });

  it('compares names ignoring case and spacing', () => {
    expect(prompts([scan({ owner: ' carmen ' })], [row()])).toHaveLength(1);
  });

  it('asks once per old machine', () => {
    expect(prompts([scan()], [row(), row({ id: 5, computerName: 'CARMEN-PC' })])).toHaveLength(2);
  });

  it('does not ask about a machine already in the stash or retired', () => {
    expect(prompts([scan()], [row({ status: 'Spare' })])).toHaveLength(0);
  });

  it('does not ask again every time a known machine is re-scanned for the same person', () => {
    const existing = [row(), row({ id: 5, computerName: 'CARMEN-PC', serialNumber: 'PC1', deviceType: 'Desktop' })];
    expect(prompts([scan({ computerName: 'CARMEN-HP', serialNumber: 'OLD1' })], existing)).toHaveLength(0);
  });

  it('asks when a known machine is re-scanned under a new person who already has one', () => {
    const existing = [row({ owner: 'Farid' }), row({ id: 5, computerName: 'CARMEN-PC', serialNumber: 'PC1' })];
    expect(prompts([scan({ computerName: 'CARMEN-HP', serialNumber: 'OLD1' })], existing)).toHaveLength(1);
  });

  it('does not ask for a report with no owner', () => {
    expect(prompts([scan({ owner: null })], [row()])).toHaveLength(0);
  });
});

describe('unanswered', () => {
  it('counts prompts with no answer, skipping rows left out of the save', () => {
    const list = prompts([scan(), scan({ sourceFileName: 'd.txt', serialNumber: 'NEW2' })], [row()]);
    expect(unanswered(list, {})).toBe(2);
    expect(unanswered(list, { 'c.txt|4': 'keep' })).toBe(1);
    expect(unanswered(list, {}, new Set(['d.txt']))).toBe(1);
  });
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `npx vitest run src/features/devices/lifecycle/matchIncoming.test.js src/features/devices/lifecycle/replacements.test.js`
Expected: FAIL.

- [ ] **Step 3: Implement**

Create `src/features/devices/lifecycle/matchIncoming.js`:

```js
import { normaliseSerial } from '../parse/placeholders.js';
import { statusOf, SPARE, RETIRED } from './status.js';

/**
 * Which register row an incoming report is about. Serial first, because the
 * computer name is the one thing IT changes when a machine changes hands;
 * name second, because every report scanned before 2026-10-05 has no serial.
 */
export const MATCH = {
  NEW: 'new',
  SAME: 'same',
  RENAMED: 'renamed',
  NAME_REUSED: 'nameReused',
  DUPLICATE_SERIAL: 'duplicateSerial',
};

const nameKey = (name) => String(name ?? '').trim().toLowerCase();
const serialKey = (serial) => (serial ? normaliseSerial(serial) : null);

export function matchIncoming(incoming, existing) {
  const bySerial = new Map();
  const byName = new Map();
  for (const row of existing) {
    const serial = serialKey(row.serialNumber);
    if (serial) bySerial.set(serial, row);
    if (row.computerName) byName.set(nameKey(row.computerName), row);
  }

  const claimed = new Map();

  return incoming.map((device) => {
    const serial = serialKey(device.serialNumber);

    if (serial) {
      const first = claimed.get(serial);
      if (first) {
        return {
          device, existing: null, kind: MATCH.DUPLICATE_SERIAL,
          note: `Same serial as ${first} in this import, so only that one is saved.`,
        };
      }
      claimed.set(serial, device.computerName ?? device.sourceFileName);

      const hit = bySerial.get(serial);
      if (hit) {
        const renamed = nameKey(hit.computerName) !== nameKey(device.computerName);
        return {
          device, existing: hit, kind: renamed ? MATCH.RENAMED : MATCH.SAME,
          note: renamed ? `Renamed from ${hit.computerName}: same serial, so its history carries over.` : null,
        };
      }
    }

    const named = byName.get(nameKey(device.computerName));
    if (named) {
      const theirs = serialKey(named.serialNumber);
      if (serial && theirs && theirs !== serial) {
        return {
          device, existing: null, kind: MATCH.NAME_REUSED,
          note: `${named.computerName} was a different machine (serial ${named.serialNumber}). This one is saved as a new machine.`,
        };
      }
      return { device, existing: named, kind: MATCH.SAME, note: null };
    }

    return { device, existing: null, kind: MATCH.NEW, note: null };
  });
}

/** The line the review shows under a row, or null when there is nothing to say. */
export function noticeFor(match) {
  if (match.note) return match.note;
  if (!match.existing) return null;
  const status = statusOf(match.existing);
  if (status === SPARE) return 'Brought back from IT Stash: it goes back into use with whoever the scan names.';
  if (status === RETIRED) return 'Brought back from the Graveyard: it goes back into use with whoever the scan names.';
  return null;
}
```

Create `src/features/devices/lifecycle/replacements.js`:

```js
import { inFleet } from './status.js';
import { MATCH } from './matchIncoming.js';

/**
 * "Carmen already has CARMEN-HP. Is this a replacement?"
 *
 * Asked only when a person turns up on a machine that is NEW to them -- a
 * machine the register has never seen, or one it has seen under somebody
 * else. Re-scanning Carmen's own laptop while she also has a desktop must not
 * ask the question again on every import.
 */
export const ANSWERS = { STASH: 'stash', RETIRE: 'retire', KEEP: 'keep' };

const ownerKey = (owner) => String(owner ?? '').trim().replace(/\s+/g, ' ').toLowerCase();

export function replacementsFor(matches, existing) {
  const prompts = [];

  for (const match of matches) {
    if (match.kind === MATCH.DUPLICATE_SERIAL) continue;
    const owner = ownerKey(match.device.owner);
    if (!owner) continue;
    if (match.existing && ownerKey(match.existing.owner) === owner) continue;

    const selfId = match.existing?.id ?? null;
    for (const row of existing) {
      if (row.id === selfId || !inFleet(row) || ownerKey(row.owner) !== owner) continue;
      prompts.push({
        key: `${match.device.sourceFileName}|${row.id}`,
        sourceFileName: match.device.sourceFileName,
        incomingName: match.device.computerName,
        owner: match.device.owner,
        old: {
          id: row.id,
          computerName: row.computerName,
          deviceType: row.deviceType,
          status: row.status || 'In use',
          since: row.createdOn ?? row.importedOn ?? null,
        },
      });
    }
  }

  return prompts;
}

export function unanswered(prompts, answers, excluded = new Set()) {
  return prompts.filter((prompt) => !excluded.has(prompt.sourceFileName) && !answers[prompt.key]).length;
}
```

- [ ] **Step 4: Run them to see them pass**

Run: `npx vitest run src/features/devices/lifecycle`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/lifecycle/matchIncoming.js src/features/devices/lifecycle/matchIncoming.test.js src/features/devices/lifecycle/replacements.js src/features/devices/lifecycle/replacements.test.js
git commit -m "Match an import by serial then name, see renames and reused names, and ask when someone gets a new machine

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 9: Planning and writing an import

**Files:**
- Modify: `src/features/devices/sharepoint/syncDevices.js`
- Modify: `src/features/devices/sharepoint/syncDevices.test.js`, `src/features/devices/sharepoint/syncPhases.test.js`

**Interfaces:**
- Consumes:
  - `matchIncoming`, `MATCH` (Task 8)
  - `replacementsFor`, `ANSWERS` (Task 8)
  - `planLifecycle`, `ACTIONS` (Task 6)
  - `currentStint` (Task 5)
  - `stintWrites`, `writeStints`, `lifecycleItem` (Task 7)
  - `readAssignments` (Task 4)
  - `inFleet`, `IN_USE` (Task 3)
- Produces:
  - `planImport(incoming, existing, { stints, answers, now, recordedBy })` returns `{ matches, prompts, inserts, updates, changeRows, stintOps, retirements, skipped }`, where:
    - `inserts: [{ computerName, body }]`
    - `updates: [{ computerName, id, body }]`
    - `changeRows: [{ computerName, deviceId, fieldName, oldValue, newValue, changeType }]`
    - `stintOps: [{ computerName, deviceId: number|null, close, open }]`
    - `retirements: [{ id, computerName, body, changes }]`
    - `skipped: [{ computerName, reason }]`
  - `syncDevices({ siteUrl, token, devices, changedBy, onProgress, answers })`. Its result gains `stintsWritten`, `stintFailures`, `skipped`.
  - `planSync` is REMOVED; its tests move to `planImport`.

- [ ] **Step 1: Port the existing tests and add new ones**

In `src/features/devices/sharepoint/syncDevices.test.js`:
- Replace the import lines with:

```js
import { describe, it, expect } from 'vitest';
import { planImport } from './syncDevices.js';
```

- Replace every `planSync([` with `planImport([`.
- Replace every `, indexByName([` argument with `, [`, and its closing `]))` with `])`. For example, `planSync([device()], indexByName([existing]))` becomes `planImport([device()], [existing])`.
- Every `changeRows` expectation gains `deviceId: 7` (the existing row's id). Every `changeRows` row also now includes any `location`/`status`/`serialNumber`/`computerName` change. The fixture `device()` has none of those, so the expectations stay as they are apart from `deviceId`.

Then append:

```js
describe('planImport — lifecycle', () => {
  const NOW = Date.UTC(2026, 9, 5);
  const scan = (over) => device({ serialNumber: 'S1', location: 'F1', ...over });
  const row = (over) => ({ ...device({ serialNumber: 'S1', location: 'F1' }), id: 7, createdOn: Date.UTC(2026, 7, 21), ...over });

  it('marks a new machine In use and opens its first stint, since at least the scan', () => {
    const plan = planImport([scan()], [], { now: NOW });
    expect(plan.inserts[0].body.Status).toBe('In use');
    expect(plan.stintOps).toHaveLength(1);
    expect(plan.stintOps[0]).toMatchObject({ computerName: 'PC1', deviceId: null, close: null });
    expect(plan.stintOps[0].open).toMatchObject({ owner: 'Ali', location: 'F1', assignedOnApprox: true });
  });

  it('renames a machine found by serial and logs the rename', () => {
    const plan = planImport([scan({ computerName: 'PC9' })], [row()], { now: NOW });
    expect(plan.updates[0]).toMatchObject({ id: 7 });
    expect(plan.updates[0].body.Title).toBe('PC9');
    expect(plan.changeRows).toContainEqual(expect.objectContaining({ fieldName: 'computerName', oldValue: 'PC1', newValue: 'PC9', deviceId: 7 }));
  });

  it('never clears a location or serial an older report does not carry', () => {
    const plan = planImport([device({ installedRamGB: 16 })], [row()], { now: NOW });
    expect(plan.updates[0].body.Location).toBe('F1');
    expect(plan.updates[0].body.SerialNumber).toBe('S1');
  });

  it('records a new owner the scan reveals as a reassignment', () => {
    const plan = planImport([scan({ owner: 'Aisyah' })], [row()], { now: NOW });
    expect(plan.stintOps).toHaveLength(1);
    expect(plan.stintOps[0].close).toMatchObject({ endReason: 'Reassigned' });
    expect(plan.stintOps[0].open).toMatchObject({ owner: 'Aisyah', assignedOnApprox: false });
  });

  it('leaves a hand-typed owner and its history alone', () => {
    const plan = planImport([scan({ owner: 'Aisyah' })], [row({ manualFields: ['owner'] })], { now: NOW });
    expect(plan.stintOps).toHaveLength(0);
  });

  it('brings a retired machine back when it is scanned again', () => {
    const plan = planImport([scan()], [row({ status: 'Retired', owner: null })], { now: NOW });
    expect(plan.updates[0].body.Status).toBe('In use');
    expect(plan.stintOps[0].open).toMatchObject({ owner: 'Ali' });
  });

  it('sends the old machine to the stash when the answer says so', () => {
    const old = row({ id: 8, computerName: 'ALI-OLD', serialNumber: 'S0' });
    const incoming = scan({ computerName: 'ALI-NEW', serialNumber: 'S2', sourceFileName: 'new.txt' });
    const plan = planImport([incoming], [old], { now: NOW, answers: { 'new.txt|8': 'stash' } });
    expect(plan.retirements).toHaveLength(1);
    expect(plan.retirements[0].body).toMatchObject({ Status: 'Spare', Owner: '' });
    expect(plan.stintOps.find((op) => op.deviceId === 8).close.endReason).toBe('Replaced');
    // The new machine is a real hand-over, not a guess about the past.
    expect(plan.stintOps.find((op) => op.computerName === 'ALI-NEW').open.assignedOnApprox).toBe(false);
  });

  it('keeps both when told to', () => {
    const old = row({ id: 8, computerName: 'ALI-OLD', serialNumber: 'S0' });
    const plan = planImport([scan({ computerName: 'ALI-NEW', serialNumber: 'S2', sourceFileName: 'new.txt' })], [old], { answers: { 'new.txt|8': 'keep' } });
    expect(plan.retirements).toHaveLength(0);
  });

  it('skips the second of two reports with one serial', () => {
    const plan = planImport([scan(), scan({ computerName: 'PC2', sourceFileName: 'x.txt' })], []);
    expect(plan.inserts).toHaveLength(1);
    expect(plan.skipped[0].computerName).toBe('PC2');
  });
});
```

In `src/features/devices/sharepoint/syncPhases.test.js`, change the item-write reply `return reply({}, 201);` to `return reply({ Id: 501 }, 201);`, so a new machine's first stint has an id to point at.

- [ ] **Step 2: Run them to see them fail**

Run: `npx vitest run src/features/devices/sharepoint/syncDevices.test.js`
Expected: FAIL. `planImport` is not exported.

- [ ] **Step 3: Implement**

In `src/features/devices/sharepoint/syncDevices.js`:

1. Replace the imports with:

```js
import { spFetch, listPath, ITEM_ACCEPT } from '../../sharepoint/spClient.js';
import { provisionLists } from './provisionLists.js';
import { DEVICE_LIST_NAME, CHANGE_LIST_NAME, toListItem } from './deviceSchema.js';
import { readAllDevices } from './readDevices.js';
import { readAssignments } from './readAssignments.js';
import { diffDevice } from './diffDevice.js';
import { stintWrites, writeStints, lifecycleItem } from './writeLifecycle.js';
import { runPool, withRetry } from '../../sharepoint/writePool.js';
import { formatMYT } from '../../../utils/malaysiaTime.js';
import { matchIncoming, MATCH } from '../lifecycle/matchIncoming.js';
import { replacementsFor, ANSWERS } from '../lifecycle/replacements.js';
import { planLifecycle, ACTIONS } from '../lifecycle/planLifecycle.js';
import { currentStint } from '../lifecycle/stints.js';
import { inFleet, IN_USE } from '../lifecycle/status.js';
```

2. Keep `applyManualOverrides` unchanged.

3. Delete `planSync` and put this in its place:

```js
const ownerKey = (owner) => String(owner ?? '').trim().replace(/\s+/g, ' ').toLowerCase();

/**
 * Pure: every decision an import makes, in one place, so all of it is
 * testable without a token. Replaces planSync, which matched by name only.
 */
export function planImport(incoming, existing, {
  stints = [], answers = {}, now = Date.now(), recordedBy = '',
} = {}) {
  const matches = matchIncoming(incoming, existing);
  const prompts = replacementsFor(matches, existing);
  const inserts = [];
  const updates = [];
  const changeRows = [];
  const stintOps = [];
  const retirements = [];
  const skipped = [];

  const replacing = new Set(prompts
    .filter((prompt) => answers[prompt.key] && answers[prompt.key] !== ANSWERS.KEEP)
    .map((prompt) => prompt.sourceFileName));

  for (const match of matches) {
    const { device, existing: row, kind } = match;
    const on = device.scannedOn ?? now;
    const holder = { owner: device.owner, location: device.location ?? null, department: device.department ?? null };

    if (kind === MATCH.DUPLICATE_SERIAL) {
      skipped.push({ computerName: device.computerName, reason: match.note });
      continue;
    }

    if (!row) {
      inserts.push({
        computerName: device.computerName,
        body: toListItem({ ...device, status: IN_USE, statusChangedOn: on }),
      });
      if (device.owner) {
        stintOps.push({
          computerName: device.computerName,
          deviceId: null,
          close: null,
          open: {
            deviceId: null, computerName: device.computerName, ...holder,
            assignedOn: on, assignedOnApprox: !replacing.has(device.sourceFileName),
            endedOn: null, endReason: null, note: null, recordedBy,
          },
        });
      }
      continue;
    }

    // An older report carries no location or serial; it must not erase them.
    const carried = {
      ...device,
      location: device.location ?? row.location ?? null,
      serialNumber: device.serialNumber ?? row.serialNumber ?? null,
    };
    const resolved = applyManualOverrides(carried, row);
    const wasActive = inFleet(row);
    resolved.status = wasActive ? (row.status ?? null) : IN_USE;
    if (!wasActive) resolved.statusChangedOn = on;

    const changes = diffDevice(row, resolved);
    if (changes.length) {
      updates.push({ computerName: device.computerName, id: row.id, body: toListItem(resolved) });
      for (const change of changes) changeRows.push({ computerName: device.computerName, deviceId: row.id, ...change });
    }

    const who = { owner: resolved.owner, location: resolved.location ?? null, department: resolved.department ?? null };
    const startFor = () => ({
      deviceId: row.id, computerName: resolved.computerName, ...who,
      assignedOn: on, assignedOnApprox: false, endedOn: null, endReason: null, note: null, recordedBy,
    });

    if (wasActive && resolved.owner && ownerKey(row.owner) !== ownerKey(resolved.owner)) {
      const stint = currentStint(row, stints);
      stintOps.push({
        computerName: resolved.computerName,
        deviceId: row.id,
        close: stint ? { stint, endedOn: on, endReason: 'Reassigned', note: null } : null,
        open: startFor(),
      });
    } else if (!wasActive && resolved.owner) {
      stintOps.push({ computerName: resolved.computerName, deviceId: row.id, close: null, open: startFor() });
    }
  }

  const handled = new Set();
  for (const prompt of prompts) {
    const answer = answers[prompt.key];
    if (!answer || answer === ANSWERS.KEEP || handled.has(prompt.old.id)) continue;
    handled.add(prompt.old.id);

    const old = existing.find((row) => row.id === prompt.old.id);
    const incomingScan = incoming.find((device) => device.sourceFileName === prompt.sourceFileName);
    const plan = planLifecycle({
      device: old,
      stints,
      action: answer === ANSWERS.RETIRE ? ACTIONS.RETIRE : ACTIONS.TO_STASH,
      input: { endReason: 'Replaced', on: incomingScan?.scannedOn ?? now },
      now,
      recordedBy,
    });
    if (plan.refusal) continue;

    retirements.push({ id: old.id, computerName: old.computerName, body: lifecycleItem(plan.fields), changes: plan.changes });
    stintOps.push({ computerName: old.computerName, deviceId: old.id, close: plan.close, open: null });
  }

  return { matches, prompts, inserts, updates, changeRows, stintOps, retirements, skipped };
}
```

4. Replace `syncDevices` with:

```js
export async function syncDevices({
  siteUrl, token, devices, changedBy, onProgress, answers = {},
}) {
  const report = (phase, done = 0, total = 0) => onProgress?.({ phase, done, total });

  report('provisioning');
  const digest = await provisionLists(siteUrl, token, {
    onProgress: (done, total) => report('provisioning', done, total),
  });

  report('reading');
  const existing = await readAllDevices(siteUrl, token);
  const stints = await readAssignments(siteUrl, token);
  const plan = planImport(devices, existing, { stints, answers, recordedBy: changedBy ?? '' });

  const post = (path, body) =>
    withRetry(() => spFetch(siteUrl, path, { token, digest, method: 'POST', body, accept: ITEM_ACCEPT }));
  const merge = (id, body) =>
    withRetry(() => spFetch(siteUrl, `${itemPath(DEVICE_LIST_NAME)}(${id})`, {
      token, digest, method: 'POST', body, accept: ITEM_ACCEPT,
      // A SharePoint update is a POST wearing these two headers.
      headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
    }));

  const work = [
    ...plan.inserts.map((entry) => ({ ...entry, action: 'insert' })),
    ...plan.updates.map((entry) => ({ ...entry, action: 'update' })),
    ...plan.retirements.map((entry) => ({ ...entry, action: 'retire' })),
  ];

  report('writing', 0, work.length);
  const results = await runPool(
    work,
    async (entry) => {
      const response = entry.action === 'insert'
        ? await post(itemPath(DEVICE_LIST_NAME), entry.body)
        : await merge(entry.id, entry.body);
      if (!response.ok) throw new Error(`${response.status}: ${await response.text()}`);
      if (entry.action !== 'insert') return { id: entry.id };
      const created = await response.json().catch(() => ({}));
      return { id: created?.Id ?? created?.d?.Id ?? null };
    },
    { concurrency: 4, onProgress: (done, total) => report('writing', done, total) },
  );

  // A new machine's first stint can only point at it once SharePoint has
  // given the row an id; a machine whose row failed gets no history at all.
  const idOf = new Map();
  const landed = new Set();
  results.forEach((result, index) => {
    if (result.error) return;
    landed.add(work[index].computerName);
    if (work[index].action === 'insert') idOf.set(work[index].computerName, result.value?.id ?? null);
  });

  let stintsWritten = 0;
  let stintFailures = 0;
  for (const op of plan.stintOps) {
    if (!landed.has(op.computerName)) continue;
    const deviceId = op.deviceId ?? idOf.get(op.computerName);
    if (deviceId === null || deviceId === undefined) {
      stintFailures += 1;
      continue;
    }
    const writes = stintWrites({ close: op.close, open: op.open ? { ...op.open, deviceId } : null });
    try {
      await writeStints({ siteUrl, token, digest, writes });
      stintsWritten += writes.length;
    } catch {
      stintFailures += 1;
    }
  }

  const retirementRows = plan.retirements.flatMap((entry) => entry.changes.map((change) => ({
    computerName: entry.computerName, deviceId: entry.id, ...change,
  })));
  const changeRows = [...plan.changeRows, ...retirementRows];

  const changedOn = Date.now();
  if (changeRows.length) report('logging', 0, changeRows.length);
  const changeResults = await runPool(changeRows, async (row) => {
    const response = await post(itemPath(CHANGE_LIST_NAME), {
      Title: row.computerName,
      DeviceId: row.deviceId ?? idOf.get(row.computerName) ?? null,
      FieldName: row.fieldName,
      OldValue: row.oldValue,
      NewValue: row.newValue,
      ChangeType: row.changeType,
      ChangedOn: new Date(changedOn).toISOString(),
      ChangedOnMYT: formatMYT(changedOn, 'datetime12'),
      ChangedBy: changedBy ?? '',
    });
    if (!response.ok) throw new Error(String(response.status));
    return true;
  }, {
    concurrency: 4,
    onProgress: (done, total) => report('logging', done, total),
  });

  return {
    results: results.map((result, index) => ({
      computerName: work[index].computerName,
      action: work[index].action,
      error: result.error ? result.error.message : null,
    })),
    changeCount: changeRows.length,
    changeFailures: changeResults.filter((result) => result.error).length,
    unchanged: devices.length - plan.inserts.length - plan.updates.length - plan.skipped.length,
    stintsWritten,
    stintFailures,
    skipped: plan.skipped,
  };
}
```

`runPool` returns `{ value, error, item }` per entry. Check `writePool.js:30` before relying on `result.value`. If it returns the worker's value under a different key, use that key in the two places above.

`diffDevice` comparisons include `status`. A matched active row whose `status` is null and `resolved.status` null compare equal. A row whose status is `'In use'` while `resolved.status` is `'In use'` also compares equal.

- [ ] **Step 4: Run the device suite**

Run: `npx vitest run src/features/devices`
Expected: PASS. If a `syncPhases.test.js` assertion counts written items, it now also sees the first-stint write for `PC1` (owner Ali). Update the count with a comment saying the extra write is the new machine's first stint.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/sharepoint
git commit -m "Plan an import with serial matching, renames, scan-revealed reassignments and replacement answers

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 9a: Every button shows that it is working

**Files:**
- Create: `src/components/ui/Spinner.jsx`
- Modify: `src/components/ui/Button.jsx`, `src/components/ui/Surfaces.jsx` (`ErrorBanner`), `src/styles/shell.css`
- Modify (existing device buttons): `src/pages/DevicesPage.jsx` (Refresh), `src/pages/DeviceDetailPage.jsx` (Refresh), `src/features/devices/ui/DropZone.jsx`, `src/features/devices/ui/DeviceTable.jsx` (the confirm Save / Remove / Remove N buttons)

**Interfaces:**
- Produces:
  - `<Spinner size={14} />`: an inline, `currentColor`, `aria-hidden` spinning ring
  - `<Button loading>`: shows the spinner in place of the icon, keeps the label, is disabled, and sets `aria-busy="true"`
  - `<ErrorBanner busy>`: its Retry button shows the spinner and is disabled

**The rule (also in Global Constraints):** any button that starts a SharePoint read or write shows the spinner from the press until the work settles. Use `<Button loading>`, or `<Spinner />` inside a raw `<button>` that cannot become a `Button`. The label never changes to "Loading…": the spinner says it is working, and the label still says what. The button keeps its width, so nothing beside it moves.

- [ ] **Step 1: Write the spinner and the button state**

Create `src/components/ui/Spinner.jsx`:

```jsx
/** The one "this is working" mark. Inherits the text colour of whatever holds it. */
export default function Spinner({ size = 14 }) {
  return <span className="ui-spinner" style={{ width: size, height: size }} aria-hidden="true" />;
}
```

Replace `src/components/ui/Button.jsx` with:

```jsx
import Spinner from './Spinner';

/**
 * The one button. Variants and sizes are class names on `.ui-btn` — see
 * `src/styles/shell.css`.
 *
 * `loading` is the universal busy state: the spinner takes the icon's place,
 * the label stays so the reader still knows WHAT is happening, and the button
 * is disabled so a second press cannot send the same write twice.
 */
export default function Button({
  children,
  variant = 'primary',
  size = 'md',
  icon: Icon,
  className = '',
  loading = false,
  disabled,
  ...props
}) {
  return (
    <button
      className={`ui-btn ui-btn-${size} ui-btn-${variant}${loading ? ' ui-btn-loading' : ''} ${className}`.trim()}
      disabled={disabled || loading}
      aria-busy={loading || undefined}
      {...props}
    >
      {loading ? <Spinner size={14} /> : Icon && <Icon size={14} />}
      {children}
    </button>
  );
}
```

In `src/components/ui/Surfaces.jsx`, give `ErrorBanner` a `busy` prop. Change its signature to accept `busy = false`. Change its Retry button to:

```jsx
        <button type="button" onClick={onRetry} disabled={busy} aria-busy={busy || undefined}>
          {busy && <Spinner size={12} />} Retry
        </button>
```

Keep the button's existing label text if it differs from `Retry`, and import `Spinner from './Spinner'`.

Append to `src/styles/shell.css`. It goes in shell.css because the button is portal-wide; check first that `@keyframes spin` exists there (it is used at line ~1341). If it does not, add it.

```css
/* The universal busy mark. A ring with one side open, turning. */
.ui-spinner {
  display: inline-block;
  flex: none;
  box-sizing: border-box;
  border: 2px solid currentColor;
  border-right-color: transparent;
  border-radius: 50%;
  animation: spin 0.8s linear infinite;
  vertical-align: -2px;
}
.ui-btn-loading { cursor: progress; }
.ui-btn-loading:disabled { opacity: 1; }
@media (prefers-reduced-motion: reduce) {
  /* Still visibly working, without the spin. */
  .ui-spinner { animation: ui-spinner-pulse 1.2s ease-in-out infinite; border-right-color: currentColor; }
}
@keyframes ui-spinner-pulse { 50% { opacity: 0.35; } }
```

`.ui-btn-loading:disabled { opacity: 1; }` keeps a busy button looking pressed-and-working rather than greyed out and broken. Check that `.ui-btn:disabled` (line ~696) sets opacity. If it uses another property, mirror it.

- [ ] **Step 2: Put it on the device section's existing buttons**

- `src/pages/DevicesPage.jsx`: on the Refresh `Button`, add `loading={loading}`. `disabled={loading}` can go, since `loading` disables it.
- `src/pages/DeviceDetailPage.jsx`: the same on its Refresh `Button`.
- `src/features/devices/ui/DropZone.jsx`: on the `Button` at line ~60, replace `disabled={busy}` with `loading={busy}`.
- `src/features/devices/ui/DeviceTable.jsx`: import `Spinner from '../../../components/ui/Spinner'`. The four raw buttons that start a write all carry `disabled={busy}`: the confirm-many Remove (~278), the confirm-many confirm (~295), the row Save (~514) and the row Remove confirm (~540). Add `aria-busy={busy || undefined}` to each, and put `{busy && <Spinner size={12} />}` as their first child. Leave the Cancel / `dt-icon` buttons alone: they start nothing.

- [ ] **Step 3: Lint and build**

Run: `npm run lint && npm run build && npm test`
Expected: only the existing `ThemeContext.jsx` lint error, a passing build, and passing tests.

- [ ] **Step 4: Commit**

```bash
git add src/components/ui src/styles/shell.css src/pages/DevicesPage.jsx src/pages/DeviceDetailPage.jsx src/features/devices/ui/DropZone.jsx src/features/devices/ui/DeviceTable.jsx
git commit -m "Give every button a loading state: a spinner in place of the icon, label kept, second press blocked

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 10: The replacement question in the import review

**Files:**
- Create: `src/features/devices/ui/ReplacementPrompt.jsx`
- Modify: `src/features/devices/ui/ReviewGrid.jsx`, `src/features/devices/reviewIssues.js`, `src/features/devices/ui/SaveProgress.jsx`, `src/pages/DevicesPage.jsx`, `src/styles/devices.css`
- Test: `src/features/devices/reviewIssues.test.js` (create if absent; else append)

**Interfaces:**
- Consumes:
  - `matchIncoming`, `noticeFor` (Task 8)
  - `replacementsFor`, `unanswered`, `ANSWERS` (Task 8)
  - `locationsIn` (Task 2)
  - `syncDevices({ …, answers })` (Task 9)
- Produces:
  - `sortForReview(devices, first = new Set())`: file names in `first` sort to the top
  - `<ReplacementPrompt prompt answer onAnswer />`
  - `<ReviewGrid … prompts answers onAnswer notices />`, where `notices` is a `Map<sourceFileName, string>`

- [ ] **Step 1: Write the failing test**

Append to (or create) `src/features/devices/reviewIssues.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { sortForReview } from './reviewIssues.js';

describe('sortForReview — replacements first', () => {
  it('puts a row waiting for an answer above everything else', () => {
    const rows = [
      { computerName: 'A', owner: 'x', sourceFileName: 'a', deviceType: 'Laptop' },
      { computerName: 'B', owner: 'y', sourceFileName: 'b', deviceType: 'Laptop' },
    ];
    expect(sortForReview(rows, new Set(['b'])).map((r) => r.computerName)).toEqual(['B', 'A']);
  });
});
```

If the file is new, the first line is the import shown. If it exists, add only the `describe` block and merge the imports.

- [ ] **Step 2: Run it to see it fail**

Run: `npx vitest run src/features/devices/reviewIssues.test.js`
Expected: FAIL. `B` does not sort first.

- [ ] **Step 3: Implement**

In `src/features/devices/reviewIssues.js`, replace `sortForReview` with:

```js
/**
 * Rows waiting on a replacement answer first -- Save cannot go ahead without
 * them -- then rows with problems, then by name.
 */
export function sortForReview(devices, first = new Set()) {
  return [...devices].sort((a, b) => {
    const aFirst = first.has(a.sourceFileName) ? 1 : 0;
    const bFirst = first.has(b.sourceFileName) ? 1 : 0;
    if (aFirst !== bFirst) return bFirst - aFirst;
    const bHasProblems = issuesFor(b).length > 0 ? 1 : 0;
    const aHasProblems = issuesFor(a).length > 0 ? 1 : 0;
    if (bHasProblems !== aHasProblems) return bHasProblems - aHasProblems;
    return (a.computerName ?? '').localeCompare(b.computerName ?? '');
  });
}
```

Create `src/features/devices/ui/ReplacementPrompt.jsx`:

```jsx
import { ANSWERS } from '../lifecycle/replacements';
import { formatMYT } from '../../../utils/malaysiaTime';

const CHOICES = [
  [ANSWERS.STASH, 'Old one to IT Stash', (name) => `${name} becomes a spare, ready for someone else`],
  [ANSWERS.RETIRE, 'Retire old one', () => 'Goes to the Graveyard. Its record is kept'],
  [ANSWERS.KEEP, 'Keep both', (name, owner) => `${owner} really uses two machines`],
];

/** One "is this a replacement?" question. Nothing is preselected: the answer is the point. */
export default function ReplacementPrompt({ prompt, answer, onAnswer }) {
  const since = prompt.old.since ? ` since at least ${formatMYT(prompt.old.since, 'date')}` : '';
  return (
    <fieldset className="rp">
      <legend className="rp-legend">
        <strong>
          {prompt.owner} already has <span className="rp-code">{prompt.old.computerName}</span>
          {' '}({prompt.old.deviceType ?? 'machine'}, {String(prompt.old.status).toLowerCase()}{since}).
        </strong>
        <span>Is <span className="rp-code">{prompt.incomingName}</span> replacing it?</span>
      </legend>
      <div className="rp-choices">
        {CHOICES.map(([value, label, hint]) => (
          <label key={value} className={`rp-choice${answer === value ? ' rp-choice-on' : ''}`}>
            <input
              type="radio"
              name={`rp-${prompt.key}`}
              value={value}
              checked={answer === value}
              onChange={() => onAnswer(prompt.key, value)}
            />
            <span>
              <strong>{label}</strong>
              <span className="rp-hint">{hint(prompt.old.computerName, prompt.owner)}</span>
            </span>
          </label>
        ))}
      </div>
    </fieldset>
  );
}
```

In `src/features/devices/ui/ReviewGrid.jsx`:
- Add `import { Fragment } from 'react';` and `import ReplacementPrompt from './ReplacementPrompt';`
- Change the signature to:

```jsx
export default function ReviewGrid({
  devices, excluded, onChange, onToggleRow, prompts = [], answers = {}, onAnswer, notices = new Map(),
}) {
```

- Add `{ key: 'location', label: 'Location', editable: true },` to `COLUMNS` after `owner`.
- Wrap each row's `<tr>` in `<Fragment key={id}>`, and move `key={id}` from the `<tr>` to the Fragment. After the `</tr>`, add:

```jsx
                  {notices.get(id) && (
                    <tr className="rg-notice-row">
                      <td colSpan={COLUMNS.length + 2}>{notices.get(id)}</td>
                    </tr>
                  )}
                  {!isExcluded && prompts.filter((p) => p.sourceFileName === id).map((prompt) => (
                    <tr key={prompt.key} className="rg-prompt-row">
                      <td colSpan={COLUMNS.length + 2}>
                        <ReplacementPrompt prompt={prompt} answer={answers[prompt.key]} onAnswer={onAnswer} />
                      </td>
                    </tr>
                  ))}
```

In `src/features/devices/ui/SaveProgress.jsx`:
- Add `stintFailures, skipped` to the destructure from `state`.
- In the summary list that has `changeCount > 0 && …` (line 85), add:

```jsx
          stintFailures > 0 && `${stintFailures} owner-history entr${stintFailures === 1 ? 'y' : 'ies'} could not be written`,
          skipped?.length > 0 && `${skipped.length} skipped: ${skipped.map((s) => s.computerName).join(', ')}`,
```

In `src/pages/DevicesPage.jsx`:
- Imports:

```jsx
import { matchIncoming, noticeFor } from '../features/devices/lifecycle/matchIncoming';
import { replacementsFor, unanswered } from '../features/devices/lifecycle/replacements';
import { locationsIn } from '../features/devices/map/locations';
```

- State: `const [answers, setAnswers] = useState({});`
- In `IDLE_SAVE`, add `stintFailures: 0, skipped: [],`.
- After `merged` is defined:

```jsx
  const matches = useMemo(() => matchIncoming(merged, saved), [merged, saved]);
  const prompts = useMemo(() => replacementsFor(matches, saved), [matches, saved]);
  const notices = useMemo(() => new Map(
    matches.map((match) => [match.device.sourceFileName, noticeFor(match)]).filter(([, text]) => text),
  ), [matches]);
  const waiting = unanswered(prompts, answers, excluded);
```

- In `handleFiles`, replace `importFiles(files)` with `importFiles(files, { knownLocations: locationsIn(saved) })`. Replace `setParsed(sortForReview(result.devices));` with:

```jsx
      const asking = new Set(replacementsFor(matchIncoming(result.devices, saved), saved)
        .map((prompt) => prompt.sourceFileName));
      setParsed(sortForReview(result.devices, asking));
```

  Add `saved` to that `useCallback`'s dependency array.

- In `resetImport`, add `setAnswers({});`.
- In `handleSave`, pass `answers,` to `syncDevices({...})`. In the success `setSave`, add `stintFailures: outcome.stintFailures, skipped: outcome.skipped,`.
- In the review head, replace the Save button with:

```jsx
              <Button size="sm" disabled={included === 0 || waiting > 0} onClick={() => handleSave(null)}>
                {waiting > 0
                  ? `Save — ${waiting} replacement${waiting === 1 ? '' : 's'} need${waiting === 1 ? 's' : ''} an answer`
                  : `Save ${included} to SharePoint`}
              </Button>
```

- After the `{flagged > 0 && …}` span, add:

```jsx
              {waiting > 0 && <span className="review-flagged"> · {waiting} replacement{waiting === 1 ? '' : 's'} to answer</span>}
```

- Pass to `<ReviewGrid>`: `prompts={prompts} answers={answers} notices={notices} onAnswer={(key, value) => setAnswers((current) => ({ ...current, [key]: value }))}`.

Append to `src/styles/devices.css`:

```css
/* --- Import review: replacement question and notices ------------------- */
.rg-prompt-row > td,
.rg-notice-row > td {
  background: var(--it-accent-wash);
  border-top: 0;
  padding: 12px 16px;
}
.rg-notice-row > td { background: var(--it-brand-wash); color: var(--it-ink); font-size: 14px; }
.rp { border: 0; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 10px; }
.rp-legend { display: flex; flex-direction: column; gap: 2px; padding: 0; font-size: 15px; color: var(--it-ink); }
.rp-code { font-family: ui-monospace, 'SFMono-Regular', Menlo, monospace; }
.rp-choices { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 10px; }
.rp-choice {
  display: flex; gap: 10px; align-items: flex-start; min-height: 44px;
  background: var(--it-panel); border: 1px solid var(--it-line); border-radius: 12px;
  padding: 12px 14px; cursor: pointer;
}
.rp-choice:hover, .rp-choice-on { border-color: var(--it-brand); }
.rp-choice input { width: 18px; height: 18px; margin: 2px 0 0; }
.rp-choice > span { display: flex; flex-direction: column; gap: 2px; font-size: 14px; }
.rp-hint { color: var(--it-ink-soft); font-size: 13px; }
```

- [ ] **Step 4: Run tests, lint and build**

Run: `npx vitest run src/features/devices && npm run lint && npm run build`
Expected: tests PASS, lint shows only the `ThemeContext.jsx` error, and the build succeeds.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices src/pages/DevicesPage.jsx src/styles/devices.css
git commit -m "Ask during import whether a person's new machine replaces their old one, and hold Save until answered

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 11: The map's numbers and layout

**Files:**
- Create: `src/features/devices/map/zones.js`, `src/features/devices/map/zones.test.js`
- Create: `src/features/devices/map/mapLayout.js`, `src/features/devices/map/mapLayout.test.js`
- Create: `src/features/devices/map/mapLinks.js`, `src/features/devices/map/mapLinks.test.js`

**Interfaces:**
- Consumes: `inFleet`, `statusOf`, `SPARE`, `RETIRED` (Task 3); `cleanLocation` (Task 2); `labelOf`, `UNASSIGNED` from `deviceFilters.js`.
- Produces:
  - `PLACES = { STASH: 'stash', GRAVEYARD: 'graveyard', NO_LOCATION: 'nolocation' }`
  - `summarise(name, devices)` → `{ name, count, laptops, desktops, other, fit, critical, attention, health: [{ level, share }] }`
  - `worldTiles(devices)` → `{ locations: [summary & { code, departments, chips: [{ name, count }] }], noLocation: summary|null, stash: summary, graveyard: summary }`
  - `locationTiles(devices, code)` → `{ centre: summary, departments: summary[], unassigned: summary|null }`
  - `machinesIn(devices, { location, department, place })` → `device[]` (worst first)
  - `searchMachines(rows, query)` → `device[]`
  - `ringLayout(count, { reserveBottom })` → `{ columns: 3, rows, centre, cells, bottom }`, where cells are 1-based `{ col, row }`
  - `neighbour(cells, index, key)` → `index`
  - `connectors(centre, cells)` → `[{ x1, y1, x2, y2 }]`, in viewBox units of 100 per cell
  - `mapHref({ location, department, place } = {})` → `string`
  - `parentHref({ location, department, place })` → `string`

- [ ] **Step 1: Write the failing tests**

Create `src/features/devices/map/zones.test.js`:

```js
import { describe, it, expect } from 'vitest';
import {
  summarise, worldTiles, locationTiles, machinesIn, searchMachines, PLACES,
} from './zones.js';

const m = (over) => ({
  id: Math.random(), computerName: 'PC', owner: 'A', location: 'F1', department: 'ENGINEERING',
  deviceType: 'Laptop', fitStatus: 'Optimal', status: null, ...over,
});

describe('summarise', () => {
  it('counts laptops, desktops and anything else apart', () => {
    const s = summarise('F1', [m(), m({ deviceType: 'Desktop' }), m({ deviceType: 'Unknown' })]);
    expect(s).toMatchObject({ count: 3, laptops: 1, desktops: 1, other: 1 });
  });

  it('counts critical and needs-attention machines and their share', () => {
    const s = summarise('F1', [m({ fitStatus: 'Critical' }), m({ fitStatus: 'Needs Attention' }), m(), m()]);
    expect(s.critical).toBe(1);
    expect(s.attention).toBe(1);
    expect(s.health).toEqual([
      { level: 'Critical', share: 0.25 }, { level: 'Needs Attention', share: 0.25 }, { level: 'Optimal', share: 0.5 },
    ]);
  });
});

describe('worldTiles', () => {
  const fleet = [
    m({ location: 'F1' }), m({ location: 'F1', department: 'SALES' }), m({ location: 'f1' }),
    m({ location: 'PML', department: 'FINANCE' }),
    m({ location: null }),
    m({ status: 'Spare', location: 'F1' }),
    m({ status: 'Retired', deviceType: 'Desktop' }),
  ];
  const world = worldTiles(fleet);

  it('has one tile per location, largest first, counting only machines in use or in repair', () => {
    expect(world.locations.map((t) => [t.code, t.count])).toEqual([['F1', 3], ['PML', 1]]);
  });

  it('lists a location\'s biggest departments as chips', () => {
    expect(world.locations[0].departments).toBe(2);
    expect(world.locations[0].chips[0]).toEqual({ name: 'ENGINEERING', count: 2 });
  });

  it('shows machines with no location only when there are some', () => {
    expect(world.noLocation.count).toBe(1);
    expect(worldTiles([m()]).noLocation).toBeNull();
  });

  it('keeps the stash and the graveyard apart from every location', () => {
    expect(world.stash.count).toBe(1);
    expect(world.graveyard).toMatchObject({ count: 1, desktops: 1 });
  });
});

describe('locationTiles', () => {
  it('keeps the same department at two locations apart', () => {
    const rows = [m({ location: 'F1', department: 'FINANCE' }), m({ location: 'PML', department: 'FINANCE' })];
    expect(locationTiles(rows, 'F1').departments.map((t) => t.count)).toEqual([1]);
  });

  it('puts machines with no department in Unassigned', () => {
    const tiles = locationTiles([m({ department: null }), m()], 'F1');
    expect(tiles.unassigned.count).toBe(1);
    expect(tiles.departments.map((t) => t.name)).toEqual(['ENGINEERING']);
    expect(tiles.centre.count).toBe(2);
  });
});

describe('machinesIn', () => {
  it('lists a department worst first', () => {
    const rows = [m({ computerName: 'B' }), m({ computerName: 'A', fitStatus: 'Critical' })];
    expect(machinesIn(rows, { location: 'F1', department: 'ENGINEERING' }).map((d) => d.computerName)).toEqual(['A', 'B']);
  });

  it('lists the stash, the graveyard and machines with no location', () => {
    const rows = [m({ status: 'Spare' }), m({ status: 'Retired' }), m({ location: '' })];
    expect(machinesIn(rows, { place: PLACES.STASH })).toHaveLength(1);
    expect(machinesIn(rows, { place: PLACES.GRAVEYARD })).toHaveLength(1);
    expect(machinesIn(rows, { place: PLACES.NO_LOCATION })).toHaveLength(1);
  });
});

describe('searchMachines', () => {
  it('finds by owner, computer name or serial', () => {
    const rows = [m({ owner: 'Amir', computerName: 'X1', serialNumber: '5CG1' }), m({ owner: 'Siti', computerName: 'X2' })];
    expect(searchMachines(rows, 'amir')).toHaveLength(1);
    expect(searchMachines(rows, '5cg')).toHaveLength(1);
    expect(searchMachines(rows, '  ')).toHaveLength(2);
  });
});
```

Create `src/features/devices/map/mapLayout.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { ringLayout, neighbour, connectors } from './mapLayout.js';

const key = (c) => `${c.col},${c.row}`;

describe('ringLayout', () => {
  it.each([0, 1, 5, 20])('places %i tiles with no overlaps and the centre free', (count) => {
    const layout = ringLayout(count, { reserveBottom: true });
    const taken = [layout.centre, layout.bottom, ...layout.cells].map(key);
    expect(layout.cells).toHaveLength(count);
    expect(new Set(taken).size).toBe(taken.length);
    expect(layout.centre).toEqual({ col: 2, row: 2 });
  });

  it('fills the top row first, deterministically', () => {
    expect(ringLayout(3).cells).toEqual([{ col: 1, row: 1 }, { col: 2, row: 1 }, { col: 3, row: 1 }]);
  });

  it('keeps the bottom middle for the graveyard, below everything else', () => {
    expect(ringLayout(4, { reserveBottom: true }).bottom).toEqual({ col: 2, row: 3 });
    const big = ringLayout(20, { reserveBottom: true });
    expect(big.bottom.row).toBe(Math.max(...big.cells.map((c) => c.row)) + 1);
  });

  it('uses the bottom middle for a tile when nothing is reserved', () => {
    expect(ringLayout(8).cells.map(key)).toContain('2,3');
    expect(ringLayout(8).bottom).toBeNull();
  });
});

describe('neighbour', () => {
  const cells = [{ col: 1, row: 1 }, { col: 2, row: 1 }, { col: 3, row: 1 }, { col: 2, row: 2 }];
  it('moves right, down and left by position', () => {
    expect(neighbour(cells, 0, 'ArrowRight')).toBe(1);
    expect(neighbour(cells, 1, 'ArrowDown')).toBe(3);
    expect(neighbour(cells, 3, 'ArrowUp')).toBe(1);
  });
  it('stays put at an edge or on another key', () => {
    expect(neighbour(cells, 0, 'ArrowLeft')).toBe(0);
    expect(neighbour(cells, 0, 'a')).toBe(0);
  });
});

describe('connectors', () => {
  it('draws from the centre of the middle cell to the centre of each tile', () => {
    expect(connectors({ col: 2, row: 2 }, [{ col: 1, row: 1 }])).toEqual([{ x1: 150, y1: 150, x2: 50, y2: 50 }]);
  });
});
```

Create `src/features/devices/map/mapLinks.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { mapHref, parentHref } from './mapLinks.js';

describe('mapHref', () => {
  it('addresses every level', () => {
    expect(mapHref()).toBe('/devices?view=map');
    expect(mapHref({ location: 'F1' })).toBe('/devices?view=map&location=F1');
    expect(mapHref({ location: 'F1', department: 'HR & ADMIN' })).toBe('/devices?view=map&location=F1&department=HR+%26+ADMIN');
    expect(mapHref({ place: 'stash' })).toBe('/devices?view=map&place=stash');
  });
});

describe('parentHref', () => {
  it('goes one level out', () => {
    expect(parentHref({ location: 'F1', department: 'X' })).toBe('/devices?view=map&location=F1');
    expect(parentHref({ location: 'F1' })).toBe('/devices?view=map');
    expect(parentHref({ place: 'stash' })).toBe('/devices?view=map');
  });
});
```

- [ ] **Step 2: Run them to see them fail**

Run: `npx vitest run src/features/devices/map`
Expected: FAIL.

- [ ] **Step 3: Implement**

Create `src/features/devices/map/zones.js`:

```js
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
```

Create `src/features/devices/map/mapLayout.js`:

```js
/**
 * Where each tile sits: a centre tile (the IT Stash, or the location itself),
 * the rest ringed round it in the order given -- callers pass largest first --
 * and on the world map the Graveyard below everything. Cells are 1-based CSS
 * grid lines in three equal columns, which is what lets the connector paths
 * be drawn in plain percentages with nothing measured.
 */
const RING = [
  { col: 1, row: 1 }, { col: 2, row: 1 }, { col: 3, row: 1 },
  { col: 1, row: 2 }, { col: 3, row: 2 },
  { col: 1, row: 3 }, { col: 3, row: 3 },
];

export function ringLayout(count, { reserveBottom = false } = {}) {
  const ring = reserveBottom ? RING : [...RING, { col: 2, row: 3 }];
  const cells = ring.slice(0, count);
  for (let row = 4; cells.length < count; row += 1) {
    for (let col = 1; col <= 3 && cells.length < count; col += 1) cells.push({ col, row });
  }

  const lowest = Math.max(3, ...cells.map((cell) => cell.row));
  const bottom = reserveBottom ? { col: 2, row: lowest > 3 ? lowest + 1 : 3 } : null;
  return {
    columns: 3,
    rows: bottom ? Math.max(lowest, bottom.row) : lowest,
    centre: { col: 2, row: 2 },
    cells,
    bottom,
  };
}

const STEP = {
  ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0],
};

/** The tile an arrow key lands on: the nearest one that way, straight lines preferred. */
export function neighbour(cells, index, key) {
  const step = STEP[key];
  const from = cells[index];
  if (!step || !from) return index;

  let best = index;
  let bestScore = Infinity;
  cells.forEach((cell, i) => {
    const dx = cell.col - from.col;
    const dy = cell.row - from.row;
    const along = dx * step[0] + dy * step[1];
    if (along <= 0) return;
    const across = Math.abs(step[0] ? dy : dx);
    const score = along + across * 2;
    if (score < bestScore) {
      bestScore = score;
      best = i;
    }
  });
  return best;
}

const middle = (cell) => ({ x: (cell.col - 0.5) * 100, y: (cell.row - 0.5) * 100 });

export function connectors(centre, cells) {
  const from = middle(centre);
  return cells.map((cell) => {
    const to = middle(cell);
    return { x1: from.x, y1: from.y, x2: to.x, y2: to.y };
  });
}
```

Create `src/features/devices/map/mapLinks.js`:

```js
/** One address per level of the map, so Back and a shared link land where they should. */
export function mapHref({ location, department, place } = {}) {
  const params = new URLSearchParams({ view: 'map' });
  if (place) params.set('place', place);
  else {
    if (location) params.set('location', location);
    if (location && department) params.set('department', department);
  }
  return `/devices?${params.toString()}`;
}

export function parentHref({ location, department, place } = {}) {
  if (place) return mapHref();
  if (department) return mapHref({ location });
  return mapHref();
}
```

- [ ] **Step 4: Run them to see them pass**

Run: `npx vitest run src/features/devices/map`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/map
git commit -m "Count every map tile (laptops, desktops, health) and lay the tiles out round a centre

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 12: The map screens

**Files:**
- Modify: `src/components/ui/Icons.jsx`
- Create:
  - `src/features/devices/ui/MapTile.jsx`
  - `src/features/devices/ui/MapGrid.jsx`
  - `src/features/devices/ui/WorldMap.jsx`
  - `src/features/devices/ui/LocationView.jsx`
  - `src/features/devices/ui/DepartmentView.jsx`
  - `src/features/devices/ui/MachineCard.jsx`
  - `src/features/devices/ui/DeviceMap.jsx`
- Modify: `src/pages/DevicesPage.jsx`, `src/styles/devices.css`

**Interfaces:**
- Consumes: everything from Task 11; `inFleet` (Task 3).
- Produces:
  - `<DeviceMap devices loading params />`, which picks a level from `params` (`location`, `department`, `place`)
  - the tab `map`, the new default `view`

- [ ] **Step 1: Add the glyphs**

Append to `src/components/ui/Icons.jsx`:

```jsx
export const Monitor = make('Monitor', (
  <>
    <rect x="3" y="4" width="18" height="12" rx="2" />
    <path d="M8 20h8M12 16v4" />
  </>
));

export const Archive = make('Archive', (
  <>
    <path d="M3 7l9-4 9 4-9 4-9-4z" />
    <path d="M3 7v10l9 4 9-4V7" />
    <path d="M12 11v10" />
  </>
));

export const Tombstone = make('Tombstone', (
  <>
    <path d="M6 21V9a6 6 0 0 1 12 0v12" />
    <path d="M4 21h16" />
    <path d="M12 9v6M9.5 11.5h5" />
  </>
));

export const MapIcon = make('MapIcon', (
  <>
    <path d="M9 4L3 6v14l6-2 6 2 6-2V4l-6 2-6-2z" />
    <path d="M9 4v14M15 6v14" />
  </>
));

export const Wrench = make('Wrench', (
  <path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.5 2.5-2.4-.6-.6-2.4z" />
));
```

- [ ] **Step 2: Build the tile and the grid**

Create `src/features/devices/ui/MapTile.jsx`:

```jsx
import { Link } from 'react-router-dom';
import { Laptop, Monitor } from '../../../components/ui/Icons';

const LEVEL_CLASS = {
  Critical: 'crit', 'Needs Attention': 'attn', Moderate: 'mod', Optimal: 'ok', Unknown: 'unk',
};

/** One place on the map. The label says everything the tile shows, for a screen reader. */
export default function MapTile({
  to, eyebrow, title, summary, chips, variant = 'zone', cell, tileRef, onKeyDown, children,
}) {
  const label = [
    title, `${summary.count} machines`, `${summary.laptops} laptops`, `${summary.desktops} desktops`,
    summary.critical ? `${summary.critical} critical` : null,
  ].filter(Boolean).join(', ');
  const style = cell ? { '--col': cell.col, '--row': cell.row } : undefined;
  const body = (
    <>
      <span className="mt-head">
        <span className="mt-eyebrow">{eyebrow}</span>
        {variant === 'zone' && (summary.critical > 0
          ? <span className="mt-badge mt-badge-crit">{summary.critical} critical</span>
          : summary.count > 0 && summary.attention === 0 && <span className="mt-badge mt-badge-ok">All clear</span>)}
      </span>
      <span className="mt-title">
        <span className="mt-name">{title}</span>
        <span className="mt-count">{summary.count} machine{summary.count === 1 ? '' : 's'}</span>
      </span>
      {summary.health.length > 0 && (
        <span className="mt-health" aria-hidden="true">
          {summary.health.map((part) => (
            <span key={part.level} className={`mt-seg mt-seg-${LEVEL_CLASS[part.level]}`} style={{ width: `${part.share * 100}%` }} />
          ))}
        </span>
      )}
      <span className="mt-split">
        <span><Laptop size={18} /> <strong>{summary.laptops}</strong> laptop{summary.laptops === 1 ? '' : 's'}</span>
        <span><Monitor size={18} /> <strong>{summary.desktops}</strong> desktop{summary.desktops === 1 ? '' : 's'}</span>
        {summary.other > 0 && <span><strong>{summary.other}</strong> other</span>}
      </span>
      {chips?.length > 0 && (
        <span className="mt-chips">
          {chips.slice(0, 3).map((chip) => <span key={chip.name} className="mt-chip">{chip.name} {chip.count}</span>)}
          {chips.length > 3 && <span className="mt-chip">+{chips.length - 3} more</span>}
        </span>
      )}
      {children}
    </>
  );

  if (!to) return <div className={`mt mt-${variant}`} style={style} aria-label={label}>{body}</div>;
  return (
    <Link to={to} className={`mt mt-${variant}`} style={style} aria-label={label} ref={tileRef} onKeyDown={onKeyDown}>
      {body}
    </Link>
  );
}
```

Create `src/features/devices/ui/MapGrid.jsx`:

```jsx
import { useRef } from 'react';
import { connectors, neighbour } from '../map/mapLayout';

/**
 * The board the tiles sit on. Arrow keys move between tiles by position;
 * the handler is on each tile, never on window, so it cannot fight a text box.
 * `items` = [{ key, cell, render(tileRef, onKeyDown) }], centre and bottom included.
 */
export default function MapGrid({ layout, items }) {
  const refs = useRef([]);
  const cells = items.map((item) => item.cell);
  const linkable = items.map((item) => item.linkable !== false);

  const onKeyDown = (index) => (event) => {
    if (!event.key.startsWith('Arrow')) return;
    let next = neighbour(cells, index, event.key);
    // Skip a tile that is not a link (a location's own centre tile).
    if (!linkable[next]) next = neighbour(cells, next, event.key);
    if (next !== index && refs.current[next]) {
      event.preventDefault();
      refs.current[next].focus();
    }
  };

  const paths = connectors(layout.centre, items.filter((item) => item.cell !== layout.centre).map((item) => item.cell));

  return (
    <div className="mg-wrap">
      <div className="mg" style={{ '--rows': layout.rows }}>
        <svg className="mg-paths" viewBox={`0 0 ${layout.columns * 100} ${layout.rows * 100}`} preserveAspectRatio="none" aria-hidden="true">
          {paths.map((p) => (
            <line key={`${p.x2},${p.y2}`} x1={p.x1} y1={p.y1} x2={p.x2} y2={p.y2} vectorEffect="non-scaling-stroke" />
          ))}
        </svg>
        {items.map((item, index) => item.render((node) => { refs.current[index] = node; }, onKeyDown(index)))}
      </div>
    </div>
  );
}
```

The `ref` callback is assigned during commit, not render, so it satisfies `react-hooks/refs`. If lint still flags `refs.current[index] = node` inside the render, move the callback creation into a `useCallback` keyed by index.

- [ ] **Step 3: Build the three levels**

Create `src/features/devices/ui/WorldMap.jsx`:

```jsx
import { useMemo } from 'react';
import MapTile from './MapTile';
import MapGrid from './MapGrid';
import { Archive, Tombstone } from '../../../components/ui/Icons';
import { worldTiles, PLACES } from '../map/zones';
import { ringLayout } from '../map/mapLayout';
import { mapHref } from '../map/mapLinks';

export default function WorldMap({ devices }) {
  const world = useMemo(() => worldTiles(devices), [devices]);
  const ringed = [
    ...world.locations.map((tile) => ({ tile, to: mapHref({ location: tile.code }), eyebrow: 'Location', variant: 'zone' })),
    ...(world.noLocation ? [{ tile: world.noLocation, to: mapHref({ place: PLACES.NO_LOCATION }), eyebrow: 'Needs a location', variant: 'pending' }] : []),
  ];
  const layout = ringLayout(ringed.length, { reserveBottom: true });

  // DOM order is the phone order: stash, locations, graveyard.
  const items = [
    {
      key: 'stash',
      cell: layout.centre,
      render: (ref, onKeyDown) => (
        <MapTile key="stash" to={mapHref({ place: PLACES.STASH })} eyebrow="Hub" title="IT Stash" summary={world.stash}
          variant="hub" cell={layout.centre} tileRef={ref} onKeyDown={onKeyDown}>
          <Archive size={20} className="mt-glyph" />
        </MapTile>
      ),
    },
    ...ringed.map((entry, index) => ({
      key: entry.tile.name,
      cell: layout.cells[index],
      render: (ref, onKeyDown) => (
        <MapTile key={entry.tile.name} to={entry.to} eyebrow={entry.eyebrow} title={entry.tile.name}
          summary={entry.tile} chips={entry.tile.chips} variant={entry.variant}
          cell={layout.cells[index]} tileRef={ref} onKeyDown={onKeyDown} />
      ),
    })),
    {
      key: 'graveyard',
      cell: layout.bottom,
      render: (ref, onKeyDown) => (
        <MapTile key="graveyard" to={mapHref({ place: PLACES.GRAVEYARD })} eyebrow="Graveyard" title="Retired"
          summary={world.graveyard} variant="grave" cell={layout.bottom} tileRef={ref} onKeyDown={onKeyDown}>
          <Tombstone size={18} className="mt-glyph" />
        </MapTile>
      ),
    },
  ];

  const fleet = world.locations.reduce((sum, t) => sum + t.count, 0) + (world.noLocation?.count ?? 0);
  const laptops = world.locations.reduce((sum, t) => sum + t.laptops, 0) + (world.noLocation?.laptops ?? 0);
  const desktops = world.locations.reduce((sum, t) => sum + t.desktops, 0) + (world.noLocation?.desktops ?? 0);

  return (
    <section className="dm">
      <header className="dm-head">
        <span className="dm-eyebrow">PMW fleet</span>
        <h2 className="dm-title">Choose a location</h2>
        <p className="dm-sub">{fleet} machines in use · {laptops} laptops · {desktops} desktops</p>
      </header>
      <MapGrid layout={layout} items={items} />
      <p className="dm-keys"><kbd>← ↑ → ↓</kbd> move <kbd>Enter</kbd> go in <kbd>Esc</kbd> back out</p>
    </section>
  );
}
```

Create `src/features/devices/ui/LocationView.jsx`:

```jsx
import { useMemo } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import MapTile from './MapTile';
import MapGrid from './MapGrid';
import { locationTiles } from '../map/zones';
import { ringLayout } from '../map/mapLayout';
import { mapHref } from '../map/mapLinks';
import { personaFor } from '../derive/persona';

export default function LocationView({ devices, location }) {
  const navigate = useNavigate();
  const tiles = useMemo(() => locationTiles(devices, location), [devices, location]);
  const ringed = [...tiles.departments, ...(tiles.unassigned ? [tiles.unassigned] : [])];
  const layout = ringLayout(ringed.length);

  const items = [
    {
      key: 'centre',
      cell: layout.centre,
      linkable: false,
      render: () => (
        <MapTile key="centre" eyebrow="Location" title={tiles.centre.name} summary={tiles.centre} variant="hub" cell={layout.centre} />
      ),
    },
    ...ringed.map((tile, index) => ({
      key: tile.name,
      cell: layout.cells[index],
      render: (ref, onKeyDown) => (
        <MapTile key={tile.name} to={mapHref({ location, department: tile.name })}
          eyebrow={personaFor(tile.name === 'Unassigned' ? null : tile.name).label ?? 'Department'}
          title={tile.name} summary={tile} cell={layout.cells[index]} tileRef={ref} onKeyDown={onKeyDown} />
      ),
    })),
  ];

  return (
    <section className="dm" onKeyDown={(event) => { if (event.key === 'Escape') navigate(mapHref()); }}>
      <nav className="dm-crumbs" aria-label="Breadcrumb">
        <Link to={mapHref()}>Map</Link><span aria-hidden="true">›</span><span aria-current="page">{tiles.centre.name}</span>
      </nav>
      <header className="dm-head">
        <span className="dm-eyebrow">Location · {tiles.centre.name}</span>
        <h2 className="dm-title">Choose a department</h2>
        <p className="dm-sub">
          {tiles.departments.length} departments · {tiles.centre.count} machines · {tiles.centre.laptops} laptops · {tiles.centre.desktops} desktops
        </p>
      </header>
      {tiles.centre.count === 0
        ? <p className="dm-empty">Nothing is recorded at {tiles.centre.name} yet.</p>
        : <MapGrid layout={layout} items={items} />}
    </section>
  );
}
```

`personaFor(...)` returns a persona object. Read `persona.js` lines 1-70 for its exact label key (`label`), and fall back to `'Department'` as shown.

Create `src/features/devices/ui/MachineCard.jsx`:

```jsx
import { Link } from 'react-router-dom';
import { Laptop, Monitor } from '../../../components/ui/Icons';
import { statusOf, SPARE, RETIRED } from '../lifecycle/status';
import { formatMYT } from '../../../utils/malaysiaTime';

const TONE = { Critical: 'crit', 'Needs Attention': 'attn', Moderate: 'mod', Optimal: 'ok' };

/** Fixed full scales, so two cards side by side compare at a glance: 15th gen, 32 GB, 1 TB. */
const bars = (device) => [
  { key: 'CPU', value: device.cpuGeneration ?? '—', share: (device.cpuGenerationRank ?? 0) / 15 },
  { key: 'RAM', value: device.installedRamGB ? `${device.installedRamGB} GB` : '—', share: (device.installedRamGB ?? 0) / 32 },
  { key: 'Disk', value: device.storageTotalGB ? `${device.storageTotalGB} GB` : '—', share: (device.storageTotalGB ?? 0) / 1024 },
];

export default function MachineCard({ device }) {
  const status = statusOf(device);
  const tone = TONE[device.fitStatus] ?? 'unk';
  let who = device.owner ?? 'No owner recorded';
  if (status === SPARE) who = 'In IT Stash';
  if (status === RETIRED) who = `Retired${device.statusChangedOn ? ` ${formatMYT(device.statusChangedOn, 'date')}` : ''}`;
  const Glyph = device.deviceType === 'Desktop' ? Monitor : Laptop;

  return (
    <Link to={`/devices/${device.id}`} className={`mc mc-${tone}`} aria-label={`${who}, ${device.computerName}, ${device.fitStatus ?? 'not judged'}`}>
      <span className="mc-head">
        <span className="mc-who">
          <strong>{who}</strong>
          <span className="mc-name">{device.computerName}</span>
        </span>
        <span className="mc-type"><Glyph size={18} /> {device.deviceType ?? 'Unknown'}</span>
      </span>
      <span className="mc-bars">
        {bars(device).map((bar) => (
          <span key={bar.key} className="mc-bar">
            <span className="mc-bar-key">{bar.key}</span>
            <span className="mc-bar-track"><span className="mc-bar-fill" style={{ width: `${Math.min(1, bar.share) * 100}%` }} /></span>
            <span className="mc-bar-value">{bar.value}</span>
          </span>
        ))}
      </span>
      <span className={`mc-verdict mc-verdict-${tone}`}>{device.fitStatus ?? 'Not judged'}</span>
    </Link>
  );
}
```

Create `src/features/devices/ui/DepartmentView.jsx`:

```jsx
import { useMemo, useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import MachineCard from './MachineCard';
import { Search } from '../../../components/ui/Icons';
import { machinesIn, searchMachines, summarise, PLACES } from '../map/zones';
import { mapHref, parentHref } from '../map/mapLinks';
import { statusOf, IN_REPAIR } from '../lifecycle/status';

const PLACE_TITLE = {
  [PLACES.STASH]: 'IT Stash', [PLACES.GRAVEYARD]: 'Graveyard', [PLACES.NO_LOCATION]: 'No location yet',
};

export default function DepartmentView({ devices, location, department, place }) {
  const navigate = useNavigate();
  const [query, setQuery] = useState('');
  const rows = useMemo(() => machinesIn(devices, { location, department, place }), [devices, location, department, place]);
  const shown = useMemo(() => searchMachines(rows, query), [rows, query]);
  const summary = useMemo(() => summarise(department ?? PLACE_TITLE[place], rows), [rows, department, place]);
  const title = place ? PLACE_TITLE[place] : department;

  const groups = place
    ? [{ key: 'all', title: null, rows: shown }]
    : [
      { key: 'use', title: 'In use', rows: shown.filter((d) => statusOf(d) !== IN_REPAIR) },
      { key: 'repair', title: 'In repair', rows: shown.filter((d) => statusOf(d) === IN_REPAIR) },
    ];

  return (
    <section className="dm" onKeyDown={(event) => {
      if (event.key === 'Escape' && event.target.tagName !== 'INPUT') navigate(parentHref({ location, department, place }));
    }}>
      <nav className="dm-crumbs" aria-label="Breadcrumb">
        <Link to={mapHref()}>Map</Link>
        {location && !place && (<><span aria-hidden="true">›</span><Link to={mapHref({ location })}>{location}</Link></>)}
        <span aria-hidden="true">›</span><span aria-current="page">{title}</span>
      </nav>
      <header className="dv-level">
        <span className="dm-eyebrow">{place ? 'Place' : `${location} · department`}</span>
        <h2 className="dv-level-title">{title}</h2>
        <div className="dv-level-stats">
          <span><strong>{summary.count}</strong> machines</span>
          <span><strong>{summary.laptops}</strong> laptops</span>
          <span><strong>{summary.desktops}</strong> desktops</span>
          {!place && <span><strong>{summary.critical}</strong> critical</span>}
          {!place && <span><strong>{summary.attention}</strong> need attention</span>}
        </div>
      </header>
      <label className="dv-search">
        <Search size={18} />
        <span className="sr-only">Search {title}</span>
        <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Owner, computer name or serial" />
      </label>
      {groups.map((group) => (group.rows.length > 0 || group.key === 'use') && (
        <div key={group.key} className="dv-group">
          {group.title && <h3 className="dv-group-title">{group.title} · {group.rows.length}</h3>}
          {group.rows.length === 0
            ? <p className="dm-empty">{query ? 'Nothing matches that search.' : 'Nothing here.'}</p>
            : <div className="dv-cards">{group.rows.map((device) => <MachineCard key={device.id} device={device} />)}</div>}
        </div>
      ))}
    </section>
  );
}
```

Create `src/features/devices/ui/DeviceMap.jsx`:

```jsx
import WorldMap from './WorldMap';
import LocationView from './LocationView';
import DepartmentView from './DepartmentView';
import { EmptyState } from '../../../components/ui/Surfaces';

/** Which level of the map the address asks for. */
export default function DeviceMap({ devices, loading, params }) {
  const location = params.get('location') ?? '';
  const department = params.get('department') ?? '';
  const place = params.get('place') ?? '';

  if (loading && devices.length === 0) return <EmptyState>Loading the register…</EmptyState>;
  if (devices.length === 0) return <EmptyState>Nothing in the register yet. Open the Import tab and drop your scan reports.</EmptyState>;
  if (place) return <DepartmentView devices={devices} place={place} />;
  if (location && department) return <DepartmentView devices={devices} location={location} department={department} />;
  if (location) return <LocationView devices={devices} location={location} />;
  return <WorldMap devices={devices} />;
}
```

- [ ] **Step 4: Wire the tab**

In `src/pages/DevicesPage.jsx`:
- `import DeviceMap from '../features/devices/ui/DeviceMap';`
- Change `const view = params.get('view') ?? 'dashboard';` to `const view = params.get('view') ?? 'map';`
- Put `['map', 'Map'],` first in the tabs array.
- When a tab button is pressed, clear the map's own keys so a later Map click starts at the world. Replace `onClick={() => setParam('view', key)}` with:

```jsx
          onClick={() => setParams((current) => {
            const next = new URLSearchParams(current);
            next.set('view', key);
            ['location', 'place'].forEach((k) => next.delete(k));
            if (view === 'map') next.delete('department');
            return next;
          })}
```

- Above `{view === 'dashboard' && (`, add `{view === 'map' && <DeviceMap devices={saved} loading={loading} params={params} />}`

The map shares the `department` key with the dashboard scope. Leaving the map clears it, so the dashboard does not open pre-scoped to the last department visited on the map.

- [ ] **Step 5: Style it**

Append to `src/styles/devices.css`:

```css
/* --- Device map ---------------------------------------------------------- */
.dm { display: flex; flex-direction: column; gap: 16px; }
.dm-head { display: flex; flex-direction: column; gap: 4px; }
.dm-eyebrow { font-size: 12px; font-weight: 700; letter-spacing: .14em; text-transform: uppercase; color: var(--it-brand); }
.dm-title { margin: 0; font-size: 26px; font-weight: 800; letter-spacing: -0.01em; color: var(--it-ink); }
.dm-sub, .dm-keys, .dm-empty { margin: 0; font-size: 14px; color: var(--it-ink-soft); }
.dm-keys kbd {
  font: inherit; font-weight: 600; border: 1px solid var(--it-line); border-bottom-width: 2px;
  border-radius: 6px; padding: 1px 6px; background: var(--it-panel); margin: 0 4px 0 12px;
}
.dm-keys kbd:first-child { margin-left: 0; }
.dm-crumbs { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; font-size: 14px; color: var(--it-ink-soft); }
.dm-crumbs a { color: var(--it-brand); font-weight: 600; }
.dm-crumbs [aria-current] { color: var(--it-ink); font-weight: 600; }

.mg-wrap {
  border: 1px solid var(--it-line); border-radius: 16px; padding: 20px;
  background-color: var(--it-canvas);
  background-image: radial-gradient(var(--it-line) 1px, transparent 1px);
  background-size: 22px 22px;
}
.mg { position: relative; display: flex; flex-direction: column; gap: 12px; }
.mg-paths { display: none; }
@media (min-width: 768px) {
  .mg { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); grid-auto-rows: minmax(170px, auto); gap: 28px; }
  .mg > .mt { grid-column: var(--col); grid-row: var(--row); }
  .mg-paths { display: block; position: absolute; inset: 0; width: 100%; height: 100%; pointer-events: none; }
  .mg-paths line { stroke: var(--it-line); stroke-width: 3; stroke-dasharray: 2 10; stroke-linecap: round; }
}

.mt {
  position: relative; z-index: 1; display: flex; flex-direction: column; gap: 8px;
  min-height: 44px; padding: 14px 16px; border-radius: 14px; text-decoration: none; color: var(--it-ink);
  background: var(--it-panel); border: 1px solid var(--it-line); box-shadow: var(--it-card-shadow);
  transition: transform 140ms cubic-bezier(.2,0,.2,1), box-shadow 140ms cubic-bezier(.2,0,.2,1);
}
a.mt:hover { transform: translateY(-2px); }
a.mt:focus-visible { outline: none; box-shadow: 0 0 0 3px var(--it-brand-line); }
.mt-hub { background: linear-gradient(150deg, var(--it-brand) 0%, var(--it-brand-deep) 100%); color: var(--it-on-brand); border-color: transparent; }
.mt-hub .mt-eyebrow, .mt-hub .mt-count, .mt-hub .mt-split { color: var(--it-on-brand-dim); }
.mt-grave { background: var(--it-thead); border-style: dashed; }
.mt-pending { border-style: dashed; }
.mt-head { display: flex; justify-content: space-between; align-items: center; gap: 8px; }
.mt-eyebrow { font-size: 11px; font-weight: 700; letter-spacing: .14em; text-transform: uppercase; color: var(--it-ink-soft); }
.mt-title { display: flex; justify-content: space-between; align-items: baseline; gap: 8px; }
.mt-name { font-size: 22px; font-weight: 800; letter-spacing: -0.01em; }
.mt-count { font-size: 13px; color: var(--it-ink-soft); }
.mt-health { display: flex; height: 8px; border-radius: 999px; overflow: hidden; background: var(--it-line); }
.mt-seg { display: block; height: 100%; }
.mt-seg-crit { background: var(--it-danger); }
.mt-seg-attn { background: var(--it-accent); }
.mt-seg-mod { background: var(--it-brand-mid); }
.mt-seg-ok { background: var(--it-good); }
.mt-seg-unk { background: var(--it-ink-soft); }
.mt-split { display: flex; flex-wrap: wrap; gap: 14px; font-size: 14px; color: var(--it-ink-soft); }
.mt-split > span { display: inline-flex; align-items: center; gap: 6px; }
.mt-split strong { color: inherit; }
.mt:not(.mt-hub) .mt-split strong { color: var(--it-ink); }
.mt-chips { display: flex; flex-wrap: wrap; gap: 6px; margin-top: auto; }
.mt-chip { font-size: 12px; padding: 3px 8px; border-radius: 6px; background: var(--it-brand-wash); color: var(--it-ink); }
.mt-badge { font-size: 12px; font-weight: 700; padding: 2px 8px; border-radius: 999px; }
.mt-badge-crit { background: var(--it-danger-wash); color: var(--it-danger); }
.mt-badge-ok { background: var(--it-good-wash); color: var(--it-good); }
.mt-glyph { position: absolute; top: 14px; right: 16px; }

.dv-level {
  display: flex; flex-direction: column; gap: 8px; padding: 20px 24px; border-radius: 16px;
  background: linear-gradient(150deg, var(--it-brand) 0%, var(--it-brand-deep) 100%); color: var(--it-on-brand);
}
.dv-level .dm-eyebrow { color: var(--it-on-brand-dim); }
.dv-level-title { margin: 0; font-size: 34px; font-weight: 800; letter-spacing: -0.02em; }
.dv-level-stats { display: flex; flex-wrap: wrap; gap: 10px; }
.dv-level-stats > span { background: rgba(255,255,255,.14); border-radius: 10px; padding: 8px 12px; font-size: 13px; }
.dv-level-stats strong { font-size: 18px; margin-right: 4px; }
.dv-search {
  display: flex; align-items: center; gap: 8px; max-width: 440px; height: 44px; padding: 0 12px;
  border: 1px solid var(--it-line); border-radius: 10px; background: var(--it-panel); color: var(--it-ink-soft);
}
.dv-search input { flex: 1; min-width: 0; border: 0; outline: none; background: transparent; font: inherit; font-size: 16px; color: var(--it-ink); }
.dv-group { display: flex; flex-direction: column; gap: 10px; }
.dv-group-title { margin: 0; font-size: 13px; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; color: var(--it-ink-soft); }
.dv-cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(250px, 1fr)); gap: 14px; }

.mc {
  display: flex; flex-direction: column; gap: 10px; padding: 14px 16px; border-radius: 14px; text-decoration: none;
  color: var(--it-ink); background: var(--it-panel); border: 1px solid var(--it-line); border-top-width: 4px;
  box-shadow: var(--it-card-shadow);
}
.mc:focus-visible { outline: none; box-shadow: 0 0 0 3px var(--it-brand-line); }
.mc-crit { border-top-color: var(--it-danger); }
.mc-attn { border-top-color: var(--it-accent); }
.mc-mod { border-top-color: var(--it-brand-mid); }
.mc-ok { border-top-color: var(--it-good); }
.mc-head { display: flex; justify-content: space-between; gap: 8px; }
.mc-who { display: flex; flex-direction: column; gap: 2px; min-width: 0; }
.mc-name { font-family: ui-monospace, 'SFMono-Regular', Menlo, monospace; font-size: 13px; color: var(--it-ink-soft); }
.mc-type { display: inline-flex; align-items: center; gap: 4px; font-size: 12px; color: var(--it-ink-soft); }
.mc-bars { display: flex; flex-direction: column; gap: 6px; }
.mc-bar { display: grid; grid-template-columns: 36px 1fr 84px; gap: 8px; align-items: center; font-size: 12px; }
.mc-bar-key { font-weight: 700; color: var(--it-ink-soft); }
.mc-bar-track { display: block; height: 7px; border-radius: 999px; background: var(--it-line); overflow: hidden; }
.mc-bar-fill { display: block; height: 100%; border-radius: 999px; background: currentColor; }
.mc-crit .mc-bar-fill { color: var(--it-danger); }
.mc-attn .mc-bar-fill { color: var(--it-accent); }
.mc-mod .mc-bar-fill { color: var(--it-brand-mid); }
.mc-ok .mc-bar-fill { color: var(--it-good); }
.mc-bar-value { text-align: right; }
.mc-verdict { align-self: flex-start; font-size: 12px; font-weight: 700; padding: 3px 9px; border-radius: 999px; background: var(--it-canvas); }
.mc-verdict-crit { background: var(--it-danger-wash); color: var(--it-danger); }
.mc-verdict-attn { background: var(--it-accent-wash); color: var(--it-ink); }
.mc-verdict-ok { background: var(--it-good-wash); color: var(--it-good); }

@media (prefers-reduced-motion: reduce) {
  .mt, a.mt:hover { transition: none; transform: none; }
}
```

If `.sr-only` is not defined globally, grep `src/` for it before relying on it. `ReviewGrid.jsx` already uses it, so it exists.

- [ ] **Step 6: Lint, build, and look at it**

Run: `npm run lint && npm run build && npx vitest run src/features/devices`
Expected: lint shows only the existing `ThemeContext.jsx` error, the build passes, and tests pass.

Then start the dev server through the preview tool (`.claude/launch.json` → `npm run dev`, port 5173) and open `/devices`. Without a sign-in it redirects to `/login`. That is fine: the check is that `/login` renders and the console shows no ReferenceError. A missing icon import blanks every page.

- [ ] **Step 7: Commit**

```bash
git add src/components/ui/Icons.jsx src/features/devices/ui src/pages/DevicesPage.jsx src/styles/devices.css
git commit -m "Add the Map tab: locations round the IT Stash, then departments, then machine cards, with laptop and desktop counts

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 13: The machine page — status, actions, owner and spec history

**Files:**
- Create: `src/features/devices/lifecycle/specHistory.js`, `src/features/devices/lifecycle/specHistory.test.js`
- Create: `src/features/devices/sharepoint/readHistory.js`
- Create: `src/features/devices/useDeviceHistory.js`
- Create: `src/features/devices/ui/LifecycleActions.jsx`, `src/features/devices/ui/AssignDialog.jsx`, `src/features/devices/ui/OwnerHistory.jsx`, `src/features/devices/ui/SpecHistory.jsx`
- Modify: `src/pages/DeviceDetailPage.jsx`, `src/features/devices/fieldGroups.js`, `src/styles/devices.css`

**Interfaces:**
- Consumes:
  - `timelineOf` (Task 5)
  - `actionsFor`, `ACTIONS` (Task 6)
  - `performLifecycle`, `retryStints` (Task 7)
  - `readAssignments` (Task 4)
  - `locationsIn` (Task 2)
  - `statusOf` (Task 3)
- Produces:
  - `specHistory(changes) → [{ day, dayLabel, rows: [{ fieldName, label, oldValue, newValue, rename }] }]`
  - `readChanges(siteUrl, token, device) → change[]`
  - `useDeviceHistory(device) → { stints, changes, loading, error, reload }`

- [ ] **Step 1: Write the failing test**

Create `src/features/devices/lifecycle/specHistory.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { specHistory } from './specHistory.js';

const at = (d, h = 4) => Date.UTC(2026, 2, d, h);

describe('specHistory', () => {
  const rows = specHistory([
    { fieldName: 'installedRamGB', oldValue: '8', newValue: '16', changedOn: at(3) },
    { fieldName: 'ramSlotsUsed', oldValue: '1', newValue: '2', changedOn: at(3, 5) },
    { fieldName: 'computerName', oldValue: 'FARID-HP', newValue: 'AMIR-HP', changedOn: at(1) },
    { fieldName: 'owner', oldValue: 'Farid', newValue: 'Amir', changedOn: at(1) },
    { fieldName: 'status', oldValue: 'In use', newValue: 'Spare', changedOn: at(1) },
  ]);

  it('groups by Malaysia day, newest first', () => {
    expect(rows.map((g) => g.dayLabel)).toEqual(['03/03/2026', '01/03/2026']);
    expect(rows[0].rows.map((r) => r.fieldName)).toEqual(['installedRamGB', 'ramSlotsUsed']);
  });

  it('leaves owner, department, location and status to the owner history', () => {
    expect(rows[1].rows.map((r) => r.fieldName)).toEqual(['computerName']);
  });

  it('marks a rename', () => {
    expect(rows[1].rows[0]).toMatchObject({ rename: true, oldValue: 'FARID-HP', newValue: 'AMIR-HP' });
  });

  it('labels fields the way the device page does', () => {
    expect(rows[0].rows[0].label).toBe('Installed RAM (GB)');
  });
});
```

- [ ] **Step 2: Run it to see it fail**

Run: `npx vitest run src/features/devices/lifecycle/specHistory.test.js`
Expected: FAIL.

- [ ] **Step 3: Implement the pure part and the reads**

Create `src/features/devices/lifecycle/specHistory.js`:

```js
import { labelFor } from '../fieldGroups.js';
import { formatMYT } from '../../../utils/malaysiaTime.js';

/** These belong to the owner history; showing them twice would make one move look like two. */
const OWNER_FIELDS = new Set(['owner', 'department', 'location', 'status', 'device']);

export function specHistory(changes = []) {
  const byDay = new Map();
  const sorted = [...changes]
    .filter((change) => !OWNER_FIELDS.has(change.fieldName) && typeof change.changedOn === 'number')
    .sort((a, b) => b.changedOn - a.changedOn);

  for (const change of sorted) {
    const dayLabel = formatMYT(change.changedOn, 'date');
    if (!byDay.has(dayLabel)) byDay.set(dayLabel, { day: change.changedOn, dayLabel, rows: [] });
    byDay.get(dayLabel).rows.push({
      fieldName: change.fieldName,
      label: change.fieldName === 'computerName' ? 'Renamed' : labelFor(change.fieldName),
      oldValue: change.oldValue,
      newValue: change.newValue,
      rename: change.fieldName === 'computerName',
    });
  }

  // Oldest-first within a day reads as the order things happened.
  return [...byDay.values()].map((group) => ({ ...group, rows: group.rows.reverse() }));
}
```

Check: `labelFor('installedRamGB')` returns the schema Title `Installed RAM (GB)`. If `fieldGroups.js` imports anything with React, move `labelFor` usage behind a plain map instead. It does not: it only imports `deviceSchema.js`.

Create `src/features/devices/sharepoint/readHistory.js`:

```js
import { spFetch, listPath } from '../../sharepoint/spClient.js';
import { CHANGE_LIST_NAME } from './deviceSchema.js';

/**
 * One machine's change rows. Rows written since 2026-10-05 carry DeviceId;
 * older rows only the computer name, so both are asked for. A machine renamed
 * BEFORE DeviceId existed loses its pre-rename rows from this view -- they are
 * still in the list, under the old name.
 *
 * A filter on a non-indexed column fails once the list passes SharePoint's
 * 5,000-item view threshold. If that day comes, index DeviceId and Title on
 * IT Device Changes in the list settings.
 */
export async function readChanges(siteUrl, token, device) {
  const name = String(device.computerName ?? '').replace(/'/g, "''");
  const filter = encodeURIComponent(`DeviceId eq ${Number(device.id)} or Title eq '${name}'`);
  let url = `${siteUrl}${listPath(CHANGE_LIST_NAME)}/items?$top=500&$filter=${filter}`;
  const rows = [];

  while (url) {
    const response = await spFetch('', url, { token });
    if (response.status === 404) return [];
    if (!response.ok) throw new Error(`Could not read the change history (${response.status})`);
    const data = await response.json();
    rows.push(...(data.d?.results ?? []));
    url = data.d?.__next ?? null;
  }

  return rows.map((row) => ({
    fieldName: row.FieldName,
    oldValue: row.OldValue ?? '',
    newValue: row.NewValue ?? '',
    changeType: row.ChangeType,
    changedBy: row.ChangedBy ?? '',
    changedOn: row.ChangedOn ? new Date(row.ChangedOn).getTime() : null,
  }));
}
```

Create `src/features/devices/useDeviceHistory.js`:

```js
import { useCallback, useEffect, useState } from 'react';
import { useSharePointToken } from '../../hooks/useRequests';
import { readAssignments } from './sharepoint/readAssignments';
import { readChanges } from './sharepoint/readHistory';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

/** Who had this machine and what changed on it. Keyed on the id, so a rename does not re-fetch. */
export function useDeviceHistory(device) {
  const getToken = useSharePointToken();
  const [state, setState] = useState({ stints: [], changes: [], loading: true, error: '' });
  const [nonce, setNonce] = useState(0);
  const reload = useCallback(() => setNonce((n) => n + 1), []);
  const id = device?.id ?? null;
  const name = device?.computerName ?? '';

  useEffect(() => {
    if (id === null) return undefined;
    let cancelled = false;
    (async () => {
      try {
        const tokenRes = await getToken();
        const [stints, changes] = await Promise.all([
          readAssignments(SHAREPOINT_SITE_URL, tokenRes.accessToken, { deviceId: id }),
          readChanges(SHAREPOINT_SITE_URL, tokenRes.accessToken, { id, computerName: name }),
        ]);
        if (!cancelled) setState({ stints, changes, loading: false, error: '' });
      } catch (failure) {
        if (!cancelled) setState((current) => ({ ...current, loading: false, error: failure.message }));
      }
    })();
    return () => { cancelled = true; };
  }, [id, name, getToken, nonce]);

  return { ...state, reload };
}
```

`setState` happens inside the async callback, not synchronously in the effect body, so `react-hooks/set-state-in-effect` does not apply. If lint disagrees, compare with how `useDevices.js` does the same thing and mirror it exactly.

- [ ] **Step 4: Build the components**

Create `src/features/devices/ui/AssignDialog.jsx`:

```jsx
import { useEffect, useId, useRef, useState } from 'react';
import Button from '../../../components/ui/Button';
import { parseFormDate } from '../../forms/toChecklistItem';

const today = () => {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
};

/**
 * Who the machine goes to. Owner, location and department are suggested from
 * the register; anything else typed is taken as written. Escape closes it and
 * is stopped here, so it cannot also unwind the page behind.
 */
export default function AssignDialog({
  title, device, owners, locations, departments, onSubmit, onCancel, busy,
}) {
  const id = useId();
  const first = useRef(null);
  const [values, setValues] = useState({
    owner: '', location: device.location ?? '', department: device.department ?? '', on: today(), note: '',
  });
  const set = (key) => (event) => setValues((current) => ({ ...current, [key]: event.target.value }));

  useEffect(() => { first.current?.focus(); }, []);

  const submit = (event) => {
    event.preventDefault();
    onSubmit({ ...values, on: parseFormDate(values.on) ?? Date.now() });
  };

  return (
    <div className="ui-confirm" role="dialog" aria-modal="true" aria-labelledby={`${id}-t`}
      onKeyDown={(event) => { if (event.key === 'Escape') { event.stopPropagation(); onCancel(); } }}>
      <button type="button" className="ui-confirm-back" aria-label="Cancel" onClick={onCancel} />
      <form className="ui-confirm-box ad" onSubmit={submit}>
        <h2 id={`${id}-t`} className="ui-confirm-title">{title}</h2>
        {device.owner && <p className="ad-now">{device.owner} has it now.</p>}
        <label className="ad-field">
          <span>New owner</span>
          <input ref={first} required list={`${id}-owners`} value={values.owner} onChange={set('owner')} />
          <datalist id={`${id}-owners`}>{owners.map((o) => <option key={o} value={o} />)}</datalist>
        </label>
        <div className="ad-row">
          <label className="ad-field">
            <span>Location</span>
            <input list={`${id}-locs`} value={values.location} onChange={set('location')} />
            <datalist id={`${id}-locs`}>{locations.map((l) => <option key={l} value={l} />)}</datalist>
          </label>
          <label className="ad-field">
            <span>Department</span>
            <input list={`${id}-depts`} value={values.department} onChange={set('department')} />
            <datalist id={`${id}-depts`}>{departments.map((d) => <option key={d} value={d} />)}</datalist>
          </label>
          <label className="ad-field">
            <span>From</span>
            <input type="date" value={values.on} onChange={set('on')} />
          </label>
        </div>
        <label className="ad-field">
          <span>Note <em>(optional)</em></span>
          <textarea rows={2} value={values.note} onChange={set('note')} />
        </label>
        <p className="ad-summary">
          {device.owner ? `${device.owner}'s time with this machine ends` : 'It goes'} on the date above
          {values.owner ? `, and ${values.owner}'s begins.` : '.'} The specs and their history stay with the machine.
        </p>
        <div className="ui-confirm-actions">
          <Button variant="secondary" type="button" onClick={onCancel}>Cancel</Button>
          <Button type="submit" loading={busy} disabled={!values.owner.trim()}>{title.startsWith('Bring') ? 'Bring back' : 'Change owner'}</Button>
        </div>
      </form>
    </div>
  );
}
```

Check that `Button` forwards `type`. Read `src/components/ui/Button.jsx`. If it hard-codes `type="button"`, pass `onClick={submit}` on the submit button instead of relying on form submit.

Create `src/features/devices/ui/LifecycleActions.jsx`:

```jsx
import { useState } from 'react';
import Button from '../../../components/ui/Button';
import AssignDialog from './AssignDialog';
import { ACTIONS, actionsFor } from '../lifecycle/planLifecycle';
import { useConfirm } from '../../../components/ui/useConfirm';

const LABEL = {
  [ACTIONS.CHANGE_OWNER]: 'Change owner',
  [ACTIONS.TO_REPAIR]: 'To repair',
  [ACTIONS.BACK_IN_USE]: 'Back in use',
  [ACTIONS.TO_STASH]: 'To IT Stash',
  [ACTIONS.RETIRE]: 'Retire…',
  [ACTIONS.ASSIGN]: 'Assign / bring back',
};

/** The buttons a machine's status allows, and the one dialog they share. */
export default function LifecycleActions({
  device, owners, locations, departments, onAction, busy,
}) {
  const [dialog, setDialog] = useState(null);
  // Which button started the work, so only THAT one spins while `busy`.
  const [pressed, setPressed] = useState(null);
  const { ask, dialog: confirm } = useConfirm();

  const press = async (action) => {
    if (action === ACTIONS.CHANGE_OWNER || action === ACTIONS.ASSIGN) {
      setPressed(action);
      setDialog(action);
      return;
    }
    setPressed(action);
    if (action === ACTIONS.RETIRE) {
      const yes = await ask({
        title: `Retire ${device.computerName}?`,
        body: 'It goes to the Graveyard and stops counting in the fleet. Its record and history are kept, and it can be brought back.',
        confirmLabel: 'Retire',
        cancelLabel: 'Keep it in use',
      });
      if (!yes) return;
    }
    onAction(action, {});
  };

  return (
    <div className="la">
      {actionsFor(device).map((action, index) => (
        <Button key={action} variant={index === 0 ? 'primary' : 'secondary'} size="sm"
          className={action === ACTIONS.RETIRE ? 'la-danger' : undefined}
          disabled={busy && pressed !== action} loading={busy && pressed === action}
          onClick={() => press(action)}>
          {LABEL[action]}
        </Button>
      ))}
      {dialog && (
        <AssignDialog
          title={dialog === ACTIONS.ASSIGN ? `Bring back ${device.computerName}` : `Change owner of ${device.computerName}`}
          device={device} owners={owners} locations={locations} departments={departments} busy={busy}
          onCancel={() => setDialog(null)}
          onSubmit={(input) => { setDialog(null); onAction(dialog, input); }}
        />
      )}
      {confirm}
    </div>
  );
}
```

Check `Button`'s accepted variants in `Button.jsx`. If `primary` is not a variant name, omit `variant` for the first button (the default is the primary fill).

Create `src/features/devices/ui/OwnerHistory.jsx`:

```jsx
import { timelineOf } from '../lifecycle/stints';
import { formatMYT } from '../../../utils/malaysiaTime';

const date = (ms) => (typeof ms === 'number' ? formatMYT(ms, 'date') : '?');

export default function OwnerHistory({ device, stints, loading }) {
  const entries = timelineOf(device, stints);
  return (
    <section className="oh card">
      <h2 className="dd-group-title">Owner history</h2>
      {loading && <p className="dm-empty">Reading the history…</p>}
      {!loading && entries.length === 0 && <p className="dm-empty">Nobody is recorded against this machine.</p>}
      <ol className="oh-list">
        {entries.map((entry) => (entry.kind === 'gap' ? (
          <li key={`gap-${entry.from}`} className="oh-item oh-gap">
            <span className="oh-dot" aria-hidden="true" />
            <span><strong>{entry.label}</strong><span className="oh-when">{date(entry.from)} – {entry.to ? date(entry.to) : 'today'}</span></span>
          </li>
        ) : (
          <li key={entry.id ?? 'legacy'} className={`oh-item${entry.current ? ' oh-current' : ''}`}>
            <span className={`oh-dot${entry.assignedOnApprox ? ' oh-dot-approx' : ''}`} aria-hidden="true" />
            <span>
              <strong>{entry.owner}</strong>
              <span className="oh-where">{[entry.location, entry.department].filter(Boolean).join(' · ')}</span>
              {entry.current && <span className="oh-now">Now</span>}
              <span className="oh-when">
                {entry.assignedOnApprox ? 'since at least ' : ''}{date(entry.assignedOn)} – {entry.endedOn ? date(entry.endedOn) : 'today'}
              </span>
              {(entry.endReason || entry.note) && (
                <span className="oh-why">{[entry.endReason, entry.note && `"${entry.note}"`].filter(Boolean).join(' · ')}</span>
              )}
            </span>
          </li>
        )))}
      </ol>
    </section>
  );
}
```

Create `src/features/devices/ui/SpecHistory.jsx`:

```jsx
import { specHistory } from '../lifecycle/specHistory';

export default function SpecHistory({ changes, loading }) {
  const days = specHistory(changes);
  return (
    <section className="sh card">
      <h2 className="dd-group-title">Spec history <span className="dd-group-hint">from each scan</span></h2>
      {loading && <p className="dm-empty">Reading the history…</p>}
      {!loading && days.length === 0 && <p className="dm-empty">No changes recorded since the first scan.</p>}
      {days.map((group) => (
        <div key={group.dayLabel} className="sh-day">
          <span className="sh-date">{group.dayLabel}</span>
          <dl className="sh-rows">
            {group.rows.map((row) => (
              <div key={`${row.fieldName}-${row.newValue}`} className={row.rename ? 'sh-rename' : undefined}>
                <dt>{row.label}</dt>
                <dd><s>{row.oldValue || '—'}</s> → <strong>{row.newValue || '—'}</strong></dd>
              </div>
            ))}
          </dl>
        </div>
      ))}
    </section>
  );
}
```

`card` is the class the `Card` surface uses. Check `Surfaces.jsx` and use its real class name, or wrap the sections in `<Card>` instead.

- [ ] **Step 5: Wire the page**

In `src/pages/DeviceDetailPage.jsx`:
- Imports:

```jsx
import { useSharePointToken } from '../hooks/useRequests';
import { useDeviceHistory } from '../features/devices/useDeviceHistory';
import LifecycleActions from '../features/devices/ui/LifecycleActions';
import OwnerHistory from '../features/devices/ui/OwnerHistory';
import SpecHistory from '../features/devices/ui/SpecHistory';
import { performLifecycle, retryStints } from '../features/devices/sharepoint/writeLifecycle';
import { statusOf } from '../features/devices/lifecycle/status';
import { locationsIn } from '../features/devices/map/locations';
import { labelOf } from '../features/devices/deviceFilters';
import { mapHref } from '../features/devices/map/mapLinks';
```

- Move the `FIT_TONE` const below all imports. It currently sits between imports, which an import/order rule may flag.
- In the component, add:

```jsx
  const getToken = useSharePointToken();
  const history = useDeviceHistory(device);
  const [acting, setActing] = useState(false);
  const [actionError, setActionError] = useState('');
  const [pending, setPending] = useState([]);

  const owners = useMemo(() => [...new Set(devices.map((d) => d.owner).filter(Boolean))].sort(), [devices]);
  const departments = useMemo(() => [...new Set(devices.map((d) => d.department).filter(Boolean))].sort(), [devices]);
  const locations = useMemo(() => locationsIn(devices), [devices]);

  const SITE = import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

  const act = async (action, input) => {
    setActing(true);
    setActionError('');
    try {
      const tokenRes = await getToken();
      const outcome = await performLifecycle({
        siteUrl: SITE, token: tokenRes.accessToken, deviceId: device.id, action, input,
        expectedStatus: statusOf(device), recordedBy: tokenRes.account?.username ?? '',
      });
      setPending(outcome.pendingStints);
      if (outcome.pendingStints.length) setActionError('The machine moved, but its owner history could not be written.');
    } catch (failure) {
      setActionError(failure.message);
    } finally {
      setActing(false);
      reload();
      history.reload();
    }
  };

  const retry = async () => {
    setActing(true);
    try {
      const tokenRes = await getToken();
      await retryStints({ siteUrl: SITE, token: tokenRes.accessToken, writes: pending });
      setPending([]);
      setActionError('');
    } catch (failure) {
      setActionError(failure.message);
    } finally {
      setActing(false);
      history.reload();
    }
  };
```

- Inside the `device ?` branch, before `<div className="dd-summary">`, add:

```jsx
          <nav className="dm-crumbs" aria-label="Breadcrumb">
            <Link to={mapHref()}>Map</Link>
            {device.location && (<><span aria-hidden="true">›</span><Link to={mapHref({ location: device.location })}>{device.location}</Link></>)}
            {device.location && (<><span aria-hidden="true">›</span><Link to={mapHref({ location: device.location, department: labelOf(device.department) })}>{labelOf(device.department)}</Link></>)}
            <span aria-hidden="true">›</span><span aria-current="page">{device.computerName}</span>
          </nav>

          <Card className="dd-life">
            <div className="dd-life-head">
              <span className={`dd-status dd-status-${statusOf(device).replace(' ', '-').toLowerCase()}`}>{statusOf(device)}</span>
              <span className="dd-life-who">
                {device.owner
                  ? <>With <strong>{device.owner}</strong>{[device.location, device.department].filter(Boolean).map((v) => ` · ${v}`).join('')}</>
                  : 'Nobody has it'}
                {device.serialNumber && <span className="dd-life-serial">Serial {device.serialNumber}</span>}
              </span>
              <LifecycleActions device={device} owners={owners} locations={locations} departments={departments} onAction={act} busy={acting} />
            </div>
            {actionError && <ErrorBanner message={actionError} busy={acting} onRetry={pending.length ? retry : () => setActionError('')} />}
          </Card>

          <div className="dd-histories">
            <OwnerHistory device={device} stints={history.stints} loading={history.loading} />
            <SpecHistory changes={history.changes} loading={history.loading} />
          </div>
          {history.error && <ErrorBanner message={history.error} onRetry={history.reload} />}
```

- Change the empty-state link to `<Link to="/devices">Back to the map</Link>`.

In `src/features/devices/fieldGroups.js`, change the identity group's `keys` to:

```js
    keys: ['computerName', 'status', 'owner', 'ownerSource', 'location', 'department', 'deviceType', 'serialNumber', 'statusChangedOn', 'anydeskId'],
```

Append to `src/styles/devices.css`:

```css
/* --- Machine page: lifecycle and history -------------------------------- */
.dd-life-head { display: flex; flex-wrap: wrap; gap: 16px; align-items: center; }
.dd-life-who { display: flex; flex-direction: column; gap: 2px; flex: 1 1 260px; font-size: 15px; }
.dd-life-serial { font-size: 13px; color: var(--it-ink-soft); font-family: ui-monospace, 'SFMono-Regular', Menlo, monospace; }
.dd-status { font-size: 12px; font-weight: 700; padding: 4px 10px; border-radius: 999px; }
.dd-status-in-use { background: var(--it-good-wash); color: var(--it-good); }
.dd-status-in-repair { background: var(--it-accent-wash); color: var(--it-ink); }
.dd-status-spare { background: var(--it-brand-wash); color: var(--it-brand); }
.dd-status-retired { background: var(--it-thead); color: var(--it-ink-soft); }
.la { display: flex; flex-wrap: wrap; gap: 8px; }
.la .ui-btn { min-height: 44px; }
.la-danger { color: var(--it-danger) !important; }
.dd-histories { display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 16px; align-items: start; }
.oh-list { list-style: none; margin: 0; padding: 0; display: flex; flex-direction: column; gap: 14px; }
.oh-item { display: grid; grid-template-columns: 18px 1fr; gap: 10px; }
.oh-item > span:last-child { display: flex; flex-direction: column; gap: 2px; }
.oh-dot { width: 12px; height: 12px; margin-top: 4px; border-radius: 50%; background: var(--it-ink-soft); }
.oh-current .oh-dot { background: var(--it-brand); box-shadow: 0 0 0 4px var(--it-brand-wash); }
.oh-dot-approx { background: transparent; border: 2px dashed var(--it-ink-soft); }
.oh-gap .oh-dot { border-radius: 3px; background: transparent; border: 2px solid var(--it-ink-soft); }
.oh-where, .oh-when { font-size: 13px; color: var(--it-ink-soft); }
.oh-why { font-size: 13px; }
.oh-now { align-self: flex-start; font-size: 12px; font-weight: 700; padding: 1px 8px; border-radius: 999px; background: var(--it-brand-wash); color: var(--it-brand); }
.sh-day { display: flex; flex-direction: column; gap: 6px; padding: 10px 0; border-top: 1px solid var(--it-line); }
.sh-day:first-of-type { border-top: 0; }
.sh-date { font-size: 12px; font-weight: 700; letter-spacing: .1em; color: var(--it-ink-soft); }
.sh-rows { margin: 0; display: flex; flex-direction: column; gap: 4px; }
.sh-rows > div { display: grid; grid-template-columns: 160px 1fr; gap: 12px; font-size: 14px; }
.sh-rows dt { color: var(--it-ink-soft); }
.sh-rows dd { margin: 0; }
.sh-rows s { color: var(--it-ink-soft); }
.sh-rename dd { font-family: ui-monospace, 'SFMono-Regular', Menlo, monospace; }
.ad { display: flex; flex-direction: column; gap: 12px; max-width: 520px; width: 100%; }
.ad-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); gap: 12px; }
.ad-field { display: flex; flex-direction: column; gap: 6px; font-size: 13px; font-weight: 600; }
.ad-field input, .ad-field textarea {
  font: inherit; font-size: 16px; font-weight: 400; min-height: 44px; padding: 0 12px;
  border: 1px solid var(--it-line); border-radius: 10px; background: var(--it-panel); color: var(--it-ink);
}
.ad-field textarea { padding: 10px 12px; resize: vertical; }
.ad-now, .ad-summary { margin: 0; font-size: 14px; color: var(--it-ink-soft); }
.ad-summary { background: var(--it-canvas); padding: 10px 12px; border-radius: 10px; }
```

- [ ] **Step 6: Tests, lint, build**

Run: `npx vitest run src/features/devices && npm run lint && npm run build`
Expected: PASS. Lint shows only the existing error.

- [ ] **Step 7: Commit**

```bash
git add src/features/devices src/pages/DeviceDetailPage.jsx src/styles/devices.css
git commit -m "Show a machine's status, owner history and spec history, with change owner, repair, stash, retire and bring back

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 14: Register and dashboard follow the lifecycle

**Files:**
- Modify: `src/features/devices/deviceFilters.js`, `src/features/devices/ui/DeviceTable.jsx`, `src/pages/DevicesPage.jsx`
- Test: `src/features/devices/deviceFilters.test.js` (append; create if absent)

**Interfaces:**
- Consumes: `statusOf`, `inFleet`, `RETIRED`, `SPARE` (Task 3); `cleanLocation` (Task 2); `performLifecycle`, `ACTIONS` (Tasks 6, 7); `mapHref`, `PLACES` (Task 11).
- Produces: filter keys `status` and `location`. The dashboard counts only `inFleet` rows.

- [ ] **Step 1: Write the failing test**

Append to (or create) `src/features/devices/deviceFilters.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { applyFilters } from './deviceFilters.js';

describe('status and location filters', () => {
  const rows = [
    { computerName: 'A', status: null, location: 'F1' },
    { computerName: 'B', status: 'Retired', location: 'f1' },
    { computerName: 'C', status: 'Spare', location: null },
  ];

  it('filters by status, reading a blank as In use', () => {
    expect(applyFilters(rows, { status: 'In use' }).map((r) => r.computerName)).toEqual(['A']);
    expect(applyFilters(rows, { status: 'Retired' }).map((r) => r.computerName)).toEqual(['B']);
  });

  it('filters by location code, and by Unassigned for none', () => {
    expect(applyFilters(rows, { location: 'F1' }).map((r) => r.computerName)).toEqual(['A', 'B']);
    expect(applyFilters(rows, { location: 'Unassigned' }).map((r) => r.computerName)).toEqual(['C']);
  });
});
```

- [ ] **Step 2: Run it to see it fail**

Run: `npx vitest run src/features/devices/deviceFilters.test.js`
Expected: FAIL.

- [ ] **Step 3: Implement**

In `src/features/devices/deviceFilters.js`:
- Add these imports at the top:

```js
import { statusOf } from './lifecycle/status.js';
import { cleanLocation } from './map/locations.js';
```

- Add to `MATCHERS`:

```js
  status: (device, value) => statusOf(device) === value,
  location: (device, value) => labelOf(cleanLocation(device.location)) === value,
```

In `src/features/devices/ui/DeviceTable.jsx`, change `LEAD_KEYS` to start with `'computerName', 'status', 'owner', 'location', 'department',` and keep the rest of the list in order.

In `src/pages/DevicesPage.jsx`:
- Imports:

```jsx
import { inFleet, statusOf, STATUSES, RETIRED, SPARE } from '../features/devices/lifecycle/status';
import { performLifecycle } from '../features/devices/sharepoint/writeLifecycle';
import { ACTIONS } from '../features/devices/lifecycle/planLifecycle';
import { mapHref } from '../features/devices/map/mapLinks';
import { PLACES } from '../features/devices/map/zones';
import { Archive } from '../components/ui/Icons';
import { useNavigate } from 'react-router-dom';
```

  Merge `useNavigate` into the existing `react-router-dom` import line rather than adding a second one.

- Add `'status', 'location'` to `FILTER_KEYS`.
- After `const { devices: saved, … } = useDevices();`, add:

```jsx
  const navigate = useNavigate();
  // Figures count the machines people are working on. A retired laptop
  // reported as a critical risk would be a figure lying about the fleet.
  const fleet = useMemo(() => saved.filter(inFleet), [saved]);
  const spareCount = useMemo(() => saved.filter((d) => statusOf(d) === SPARE).length, [saved]);
  // The register hides retired machines unless a status filter asks for them.
  const registerRows = useMemo(
    () => (filters.status ? saved : saved.filter((d) => statusOf(d) !== RETIRED)),
    [saved, filters.status],
  );
  const locationOptions = useMemo(() => locationsIn(saved), [saved]);
```

  `filters` is defined after this point in the current file. Place this block after `filters`.

- In `departments` and `scoped`, replace `saved` with `fleet` (both the `useMemo` bodies and their dependency arrays).
- After the `Office compliance` `StatusCard`, add:

```jsx
            <StatCard
              icon={Archive}
              label="In IT Stash"
              value={spareCount}
              loading={loading}
              onClick={() => navigate(mapHref({ place: PLACES.STASH }))}
            />
```

- Before `<DeviceTable`, add:

```jsx
          <div className="dv-register-scope">
            <label className="dv-scope">
              <span>Location</span>
              <select value={filters.location} onChange={(event) => setParam('location', event.target.value)}>
                <option value="">All locations</option>
                {locationOptions.map((code) => <option key={code} value={code}>{code}</option>)}
                <option value="Unassigned">No location yet</option>
              </select>
            </label>
            <label className="dv-scope">
              <span>Status</span>
              <select value={filters.status} onChange={(event) => setParam('status', event.target.value)}>
                <option value="">All but retired</option>
                {STATUSES.map((status) => <option key={status} value={status}>{status}</option>)}
              </select>
            </label>
            {!filters.status && (
              <button type="button" className="dv-linkish" onClick={() => setParam('status', RETIRED)}>
                Retired machines are hidden. Show them
              </button>
            )}
          </div>
```

- Change `<DeviceTable devices={saved}` to `<DeviceTable devices={registerRows}`.
- Replace `handleRowSave` with:

```jsx
  const OWNERSHIP = ['owner', 'location', 'department'];
  const trimmed = (value) => String(value ?? '').trim();

  /**
   * An owner, location or department typed into the register is a change of
   * hands, so it goes through the same write as the machine page's Change
   * owner -- the register and the owner history cannot disagree. Clearing the
   * owner still goes the old way: a blank hands the field back to the scan.
   */
  const handleRowSave = (device, edits) => runRowAction(async (tokenRes) => {
    const moved = OWNERSHIP.some((key) => key in edits && trimmed(edits[key]) !== trimmed(device[key]));
    const owner = trimmed('owner' in edits ? edits.owner : device.owner);
    let existing = device;
    let rest = edits;

    if (moved && owner && inFleet(device)) {
      const { plan, pendingStints } = await performLifecycle({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        deviceId: device.id,
        action: ACTIONS.CHANGE_OWNER,
        input: {
          owner,
          location: 'location' in edits ? edits.location : device.location,
          department: 'department' in edits ? edits.department : device.department,
        },
        expectedStatus: statusOf(device),
        recordedBy: tokenRes.account?.username ?? '',
      });
      // The remaining edit must start from the manual list the change just
      // wrote, or saving the device type would drop owner from it again.
      existing = { ...device, ...plan.fields };
      rest = Object.fromEntries(Object.entries(edits).filter(([key]) => !OWNERSHIP.includes(key)));
      if (pendingStints.length) {
        throw new Error('The owner changed, but its history could not be written. Open the machine to retry.');
      }
    }

    if (Object.keys(rest).length) {
      await updateDevice({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        existing,
        edits: rest,
        changedBy: tokenRes.account?.username ?? '',
      });
    }
  });
```

  `locationsIn` is already imported from Task 10.

Append to `src/styles/devices.css`:

```css
.dv-register-scope { display: flex; flex-wrap: wrap; gap: 12px; align-items: flex-end; margin-bottom: 12px; }
.dv-register-scope select { font-size: 16px; min-height: 44px; }
.dv-linkish { background: none; border: 0; padding: 0; min-height: 44px; font: inherit; font-size: 14px; color: var(--it-brand); cursor: pointer; text-decoration: underline; }
```

- [ ] **Step 4: Tests, lint, build**

Run: `npm test && npm run lint && npm run build`
Expected: all tests PASS, lint shows only the existing error, and the build succeeds.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices src/pages/DevicesPage.jsx src/styles/devices.css
git commit -m "Count only machines in use on the dashboard, filter the register by location and status, and route owner edits through the history

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 15: Write it down

**Files:**
- Modify: `AGENTS.md`

- [ ] **Step 1: Update `AGENTS.md`**

In the ROUTES table, replace the `/devices` row with:

```
| `/devices` | Device list. `?view=map` (default): locations → departments → machines, with the IT Stash and the Graveyard; `?location=`, `&department=`, `?place=stash|graveyard|nolocation`. Also `dashboard`, `register`, `import` |
| `/devices/:id` | One machine: status, change owner / repair / stash / retire / bring back, owner history, spec history, every scanned field |
```

In WHERE TO LOOK, add rows:

```
| A machine's status, and what counts in the fleet | `src/features/devices/lifecycle/status.js` |
| What a change of owner, repair, stash, retire or bring-back writes | `src/features/devices/lifecycle/planLifecycle.js`, `sharepoint/writeLifecycle.js` |
| Owner history (stints) and its list | `lifecycle/stints.js`, `sharepoint/assignmentSchema.js` |
| How an import recognises a machine (serial, then name) | `src/features/devices/lifecycle/matchIncoming.js` |
| "Is this a replacement?" during import | `lifecycle/replacements.js`, `ui/ReplacementPrompt.jsx` |
| Location out of the scan file name | `derive/deriveLocation.js`, `map/locations.js` |
| The map's numbers and layout | `map/zones.js`, `map/mapLayout.js`, `map/mapLinks.js` |
```

In CONVENTIONS, add:

```markdown
**A machine is known by its serial, and it has a life.** The scan script writes
`Serial Number: $((Get-CimInstance Win32_BIOS).SerialNumber)`; `matchIncoming`
matches on it first and on the computer name second, so renaming a PC on
reassignment keeps its history. BIOS filler (`Default string`, `To be filled by
O.E.M.`, `0`, one character repeated) is no serial at all -- matching on it
would merge every home-built desktop into one row. Every machine has a status
(`In use`, `In repair`, `Spare`, `Retired`; blank reads In use, no migration),
and only In use and In repair count in any figure. Who had it lives in
`IT Device Assignments`, one row per stint, tied by `DeviceId`; the device row's
owner is only the present. A machine nobody has touched since this landed has
no stints, and its first change writes the outgoing owner in "since at least"
its creation date. Every lifecycle write re-reads the machine first and goes
device row → stints → change log, so a failure leaves a status the history
cannot yet explain (retryable), never history describing a move that did not
happen.

**Location comes first in the file name's bracket.** `[F1 ENGINEERING] X.txt`
is location F1, department ENGINEERING; `[ENGINEERING] X.txt` still reads as it
always did. `STOCKYARDF1` and `PML GUARDHOUSE` split into location and
department. Locations are a TEXT column of upper-case codes -- built-in F1, F3,
PML plus whatever a row uses -- so a new site needs no column change. A
location typed by hand is in `ManualFields` and beats the file; an older report
with no location never clears one.
```

- [ ] **Step 2: Commit**

```bash
git add AGENTS.md
git commit -m "Document machine identity, lifecycle, owner history and locations

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 16: Final check

- [ ] **Step 1: Everything green**

Run: `npm test && npm run lint && npm run build`
Expected: all tests pass, lint reports only `ThemeContext.jsx`, and the build succeeds.

- [ ] **Step 2: The app still opens**

Start the dev server with the preview tool, using `.claude/launch.json` (create it with `npm run dev` on port 5173 if it is absent). Open `/`, and check the console has no errors. Without a Microsoft sign-in, `/devices` redirects to `/login`, which proves the shell and the new imports load. A signed-in check of the map against real rows is left to the user, and the report must say so.

- [ ] **Step 3: Report**

Summarise for the user, in plain terms:
- what the map shows
- what the scan-script line is
- the file-name convention
- that existing machines appear under "No location yet" until re-imported or set by hand
- that the first time anyone saves an import or a lifecycle action, SharePoint gains four columns and a new list
- the 5,000-item filter limit noted in `readHistory.js`

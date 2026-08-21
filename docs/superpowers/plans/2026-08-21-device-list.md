# Device List Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a `/devices` section where IT drops the plain-text machine scan reports, reviews a parsed table, and saves one SharePoint row per machine with a change log and a fleet dashboard.

**Architecture:** A pure-function pipeline — `parseReport` → `deriveDevice` → review grid → `syncDevices` — where every stage before the grid is a pure function of the previous stage's output. Parsing keys off a known-label whitelist rather than a generic `Word:` regex. SharePoint is written only after the user approves the review table; reads are paged and diffed in memory so one batch costs one read.

**Tech Stack:** React 19, Vite 8, Vitest (already configured), SharePoint REST via delegated MSAL token. No new dependencies.

**Spec:** `docs/superpowers/specs/2026-08-21-device-list-design.md` — read it before starting. Every task cites the spec sections it implements; where this plan and the spec disagree, the spec wins and you should stop and flag it.

## Global Constraints

These apply to **every** task. They come from the spec and from `AGENTS.md`, and an engineer new to this repo will get them wrong by default.

- **Pages do not gate themselves.** `src/components/AppShell.jsx` is the one auth gate. A page renders `<AppShell title subtitle actions>` plus its own body — never its own nav, theme toggle, or sign-out.
- **Any SharePoint token comes from `useSharePointToken()`** in `src/hooks/useRequests.js`. Never call `acquireTokenSilent` or `acquireTokenPopup` from a page or a feature module — the popup fallback that replaced is where the "it just stopped loading" reports came from.
- **Never call React Router's `navigate()` inside `useEffect`** — it re-renders, re-runs the effect, and loops forever. Use `window.location.replace()` in effects; `navigate()` is correct in event handlers.
- **Never export a helper from a file that also exports a component.** It drops the file out of Fast Refresh and fails `npm run lint`. Helpers go in their own module (this is why `src/utils/initials.js` exists).
- **Design tokens, exact names:** `--it-brand`, `--it-brand-deep`, `--it-brand-mid`, `--it-brand-line`, `--it-brand-wash`, `--it-canvas`, `--it-panel`, `--it-ink`, `--it-ink-soft`, `--it-line`, `--it-accent`, `--it-good`, `--it-danger`, `--it-radius`, `--it-card-shadow`. Dark mode is `[data-theme='dark']` on the root element. Never hardcode a hex value.
- **Mobile-first CSS.** `min-width` breakpoints at `640px` and `1024px` only. Honour `@media (prefers-reduced-motion: reduce)`.
- **Stylesheet order in `src/main.jsx` is load-bearing** — `index.css` → `App.css` → `shell.css` → `auth.css`. Add `devices.css` **after** `shell.css`.
- **ESM only.** `package.json` has `"type": "module"`; there is no CommonJS in this repo.
- **Tests run in `environment: 'node'`** and are matched by `include: ['src/**/*.test.js']`. Do not add jsdom, and do not name a test file `.test.jsx` — it will not be collected.
- **Commit after every task.** Branch is `feat/data-studio`.
- `npm run lint` has **pre-existing** errors in `FormPage`, `AssetChecklistPage`, `SignatureDialog` and `ThemeContext`. Those are not yours. Do not fix them, and do not let them mask new errors in files you create — check the filenames in the output.
- **Dates:** store `toISOString()`, display through `formatMYT`. Never build a Malaysian time by adding 8 hours to anything.

---

## File Structure

| File | Responsibility |
|---|---|
| `src/features/devices/parse/placeholders.js` | The junk-token table and `isPlaceholder` |
| `src/features/devices/parse/labels.js` | The 21 known labels, normalisation, `matchLabel` |
| `src/features/devices/parse/parseReport.js` | Report text → `{ fields, unknownLabels, warnings }` |
| `src/features/devices/parse/parseValues.js` | ` \| `-delimited sub-structure parsers, one per field |
| `src/features/devices/derive/deriveRam.js` | Sticks, installed vs reported, slots, upgradability |
| `src/features/devices/derive/deriveStorage.js` | Drives, totals, SSD/HDD classification |
| `src/features/devices/derive/deriveCpu.js` | Vendor, model, generation, age band |
| `src/features/devices/derive/deriveIdentity.js` | Owner chain, department, device type, computer name |
| `src/features/devices/derive/deriveHealth.js` | Windows support, antivirus status, scan completeness |
| `src/features/devices/derive/riskScore.js` | Additive score, reasons, level |
| `src/features/devices/derive/deriveDevice.js` | Orchestrates the above into one device record |
| `src/features/devices/importFiles.js` | `File[]` → device records, batch dedupe |
| `src/features/devices/sharepoint/deviceSchema.js` | Column definitions for both lists + row mapping |
| `src/features/devices/sharepoint/provisionLists.js` | Idempotent list and column creation |
| `src/features/devices/sharepoint/readDevices.js` | Paged read, row → device record |
| `src/features/devices/sharepoint/diffDevice.js` | Tracked-field diff → change rows |
| `src/features/devices/sharepoint/writePool.js` | Concurrency-limited writes with backoff |
| `src/features/devices/sharepoint/syncDevices.js` | Read → diff → write orchestration |
| `src/features/devices/stats/deviceStats.js` | Dashboard aggregations, all pure |
| `src/features/devices/useDevices.js` | The one SharePoint read for this section |
| `src/features/devices/ui/DropZone.jsx` | Drag & drop + file picker |
| `src/features/devices/ui/ReviewGrid.jsx` | Editable review table |
| `src/features/devices/ui/SaveProgress.jsx` | Progress + per-row results |
| `src/features/devices/ui/DeviceTable.jsx` | The saved register |
| `src/features/devices/ui/DeviceCharts.jsx` | The CSS charts |
| `src/features/devices/ui/Leaderboards.jsx` | The six ranked tables |
| `src/pages/DevicesPage.jsx` | Route: AppShell + stage switch |
| `src/styles/devices.css` | All styling for the section |

Layering rule, enforced by review: `parse/` imports nothing from `derive/`, `derive/` imports nothing from `sharepoint/`, and `sharepoint/` imports no React.

---

# PHASE 1 — Parsing

---

## Task 1: Placeholder tokens and the label table

Implements spec §7.2, §7.7.

**Files:**
- Create: `src/features/devices/parse/placeholders.js`
- Create: `src/features/devices/parse/labels.js`
- Test: `src/features/devices/parse/labels.test.js`

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `isPlaceholder(value: string): boolean`
  - `cleanValue(value: string): string | null` — trims, strips NBSP/zero-width, returns `null` for placeholders
  - `KNOWN_LABELS: string[]` — the 21 labels in file order
  - `matchLabel(line: string): { label: string, inline: string } | null`

- [ ] **Step 1: Write the failing test**

`src/features/devices/parse/labels.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { isPlaceholder, cleanValue } from './placeholders.js';
import { KNOWN_LABELS, matchLabel } from './labels.js';

describe('isPlaceholder', () => {
  it('treats the unset SMBIOS strings as placeholders', () => {
    expect(isPlaceholder('Manufacturer1')).toBe(true);
    expect(isPlaceholder('PartNum1')).toBe(true);
    expect(isPlaceholder('System Product Name')).toBe(true);
    expect(isPlaceholder('To Be Filled By O.E.M.')).toBe(true);
    expect(isPlaceholder('Default string')).toBe(true);
  });

  it('treats None, Unknown and blanks as placeholders, case-insensitively', () => {
    expect(isPlaceholder('None')).toBe(true);
    expect(isPlaceholder('none')).toBe(true);
    expect(isPlaceholder('Unknown')).toBe(true);
    expect(isPlaceholder('   ')).toBe(true);
  });

  it('does not treat real values as placeholders', () => {
    expect(isPlaceholder('Samsung')).toBe(false);
    expect(isPlaceholder('HP Laptop 15-fd0xxx')).toBe(false);
  });
});

describe('cleanValue', () => {
  it('trims trailing whitespace the scan writes after the processor name', () => {
    expect(cleanValue('AMD Ryzen 5 7430U with Radeon Graphics         '))
      .toBe('AMD Ryzen 5 7430U with Radeon Graphics');
  });

  it('strips non-breaking and zero-width characters', () => {
    expect(cleanValue('HP\u00a0Laptop\u200b')).toBe('HP Laptop');
  });

  it('returns null for placeholders', () => {
    expect(cleanValue('None')).toBe(null);
    expect(cleanValue('')).toBe(null);
  });
});

describe('KNOWN_LABELS', () => {
  it('has the 21 labels the scan writes', () => {
    expect(KNOWN_LABELS).toHaveLength(21);
    expect(KNOWN_LABELS[0]).toBe('Name');
    expect(KNOWN_LABELS).toContain('Email data files found Active or Inactive account');
  });
});

describe('matchLabel', () => {
  it('matches a bare label with no inline value', () => {
    expect(matchLabel('Computer Name:')).toEqual({ label: 'Computer Name', inline: '' });
  });

  it('matches an inline value with a space after the colon', () => {
    expect(matchLabel('Antivirus status: NORTON ACTIVE'))
      .toEqual({ label: 'Antivirus status', inline: 'NORTON ACTIVE' });
  });

  it('matches an inline value with no space after the colon', () => {
    expect(matchLabel('Antivirus status:NORTON NOT INSTALLED'))
      .toEqual({ label: 'Antivirus status', inline: 'NORTON NOT INSTALLED' });
  });

  it('is case-insensitive and tolerates collapsed whitespace', () => {
    expect(matchLabel('total  ram:')).toEqual({ label: 'Total RAM', inline: '' });
  });

  it('does NOT match the RAM slot summary line', () => {
    expect(matchLabel('Total Slots: 2 | Used Slots: 2')).toBe(null);
  });

  it('does NOT match a mapped drive line', () => {
    expect(matchLabel('Y: | \\\\server\\PMW\\IT')).toBe(null);
  });

  it('does NOT match a value line that happens to contain a colon', () => {
    expect(matchLabel('Wi-Fi | SSID: PMW_Group | IP: 192.168.1.170 | Dynamic')).toBe(null);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/parse/labels.test.js`
Expected: FAIL — `Failed to resolve import "./placeholders.js"`

- [ ] **Step 3: Write `placeholders.js`**

```js
/**
 * Tokens the scan writes when it could not read a real value. Storing them
 * verbatim would produce a dashboard category called "Manufacturer1", so they
 * become null at the parse boundary rather than being filtered downstream.
 */
const PLACEHOLDER_TOKENS = new Set([
  'none',
  'unknown',
  'n/a',
  'na',
  'nil',
  'manufacturer1',
  'partnum1',
  'system product name',
  'to be filled by o.e.m.',
  'default string',
  'not specified',
  '',
]);

/** Non-breaking space, zero-width space, zero-width non-joiner, BOM. */
const INVISIBLE = /[\u00a0\u200b\u200c\ufeff]/g;

export function isPlaceholder(value) {
  if (value == null) return true;
  return PLACEHOLDER_TOKENS.has(String(value).replace(INVISIBLE, ' ').trim().toLowerCase());
}

export function cleanValue(value) {
  if (value == null) return null;
  const cleaned = String(value).replace(INVISIBLE, ' ').replace(/\s+$/, '').trim();
  return isPlaceholder(cleaned) ? null : cleaned;
}
```

- [ ] **Step 4: Write `labels.js`**

```js
/**
 * The labels the scan script writes, in file order. Verified identical across
 * all 17 sample reports.
 *
 * This list is the ONLY thing allowed to open a new field. A generic /^\w+:/
 * split reads "Total Slots: 2 | Used Slots: 2" and "Y: | \\server\PMW\IT" as
 * fields, which silently moves those values out of the block they belong to.
 */
export const KNOWN_LABELS = [
  'Name',
  'Anydesk',
  'Antivirus status',
  'Remarks',
  'Computer Name',
  'Computer Model',
  'Motherboard',
  'Windows Version',
  'Processor',
  'GPU',
  'Total RAM',
  'RAM Slot Info',
  'Storage Drives',
  'Network Information',
  'Antivirus',
  'Monitor',
  'PMW Server and credentials',
  'Server folder',
  'Microsoft Office',
  'Adobe',
  'Email data files found Active or Inactive account',
];

const normalise = (s) => s.replace(/\s+/g, ' ').trim().toLowerCase();

const BY_NORMALISED = new Map(KNOWN_LABELS.map((label) => [normalise(label), label]));

/**
 * Returns the canonical label and any inline value, or null when the line is
 * not a label. Splits on the FIRST colon only, so an inline value may itself
 * contain colons.
 */
export function matchLabel(line) {
  const colon = line.indexOf(':');
  if (colon === -1) return null;

  const label = BY_NORMALISED.get(normalise(line.slice(0, colon)));
  if (!label) return null;

  return { label, inline: line.slice(colon + 1).trim() };
}
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/parse/labels.test.js`
Expected: PASS — 12 tests.

Note why `matchLabel('Total Slots: 2 | Used Slots: 2')` returns null without a
special case: `total slots` is not in `BY_NORMALISED`. The whitelist does the
work; there is no heuristic to tune.

- [ ] **Step 6: Commit**

```bash
git add src/features/devices/parse/
git commit -m "Add device report label table and placeholder tokens"
```

---

## Task 2: `parseReport` and the real-file fixtures

Implements spec §7.1, §7.3, §7.4, §7.5.

**Files:**
- Create: `src/features/devices/parse/parseReport.js`
- Create: `src/features/devices/__fixtures__/` (copies of the real reports)
- Test: `src/features/devices/parse/parseReport.test.js`

**Interfaces:**
- Consumes: `matchLabel`, `KNOWN_LABELS` from Task 1.
- Produces: `parseReport(text: string): { fields: Record<string, string[]>, unknownLabels: {label,value}[], warnings: string[], isReport: boolean }`
  - `fields` is keyed by canonical label; every value is an array of answers, possibly empty.
  - `isReport` is false when no known label was found at all.

- [ ] **Step 1: Copy the fixture files**

Copy the real reports so the tests run against real input, not idealised input:

```bash
mkdir -p src/features/devices/__fixtures__
cp "$HOME/Downloads/ASHRAF-PC_.txt" src/features/devices/__fixtures__/
cp "$HOME/Downloads/CARMEN-HP_.txt" src/features/devices/__fixtures__/
cp "$HOME/Downloads/[ENGINEERING] AMIR-HP_.txt" src/features/devices/__fixtures__/
cp "$HOME/Downloads/[QAQC FAIRUS]HPFL05_.txt" src/features/devices/__fixtures__/
cp "$HOME/Downloads/[STOCKYARDF1] DESKTOP-8SBR420_.txt" src/features/devices/__fixtures__/
cp "$HOME/Downloads/[SALES] PGCHAN-HP_.txt" src/features/devices/__fixtures__/
cp "$HOME/Downloads/[FINANCE] EVONNE-HP_.txt" src/features/devices/__fixtures__/
cp "$HOME/Downloads/PMWL034_.txt" src/features/devices/__fixtures__/
```

These eight cover every parsing edge in the set: CRLF, inline value with and
without a space, the two label look-alikes, a fully blank report, 22 duplicate
antivirus lines, a single-stick machine, and Core Ultra.

- [ ] **Step 2: Write the failing test**

`src/features/devices/parse/parseReport.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { parseReport } from './parseReport.js';

const fixture = (name) =>
  readFileSync(fileURLToPath(new URL(`../__fixtures__/${name}`, import.meta.url)), 'utf8');

describe('parseReport — structure', () => {
  const parsed = parseReport(fixture('ASHRAF-PC_.txt'));

  it('reads a single-answer field', () => {
    expect(parsed.fields['Computer Name']).toEqual(['ASHRAF-PC']);
  });

  it('reads a multi-answer field as every line', () => {
    expect(parsed.fields['GPU']).toEqual([
      'Intel(R) Iris(R) Xe Graphics',
      'VirtualMonitorDriver Device',
    ]);
  });

  it('keeps the RAM slot summary inside the RAM Slot Info block', () => {
    expect(parsed.fields['RAM Slot Info']).toEqual([
      '4 GB | DDR4 | 3200 MHz | Samsung | M471A5244CB0-CWE',
      '4 GB | DDR4 | 3200 MHz | Samsung | M471A5244CB0-CWE',
      'Total Slots: 2 | Used Slots: 2',
    ]);
  });

  it('keeps mapped drive lines inside Server folder', () => {
    expect(parsed.fields['Server folder']).toEqual([
      'Y: | \\\\server\\emdata$\\device list 2026',
      'Z: | \\\\server\\PMW\\IT',
    ]);
  });

  it('does not leak the banner text into Remarks', () => {
    expect(parsed.fields['Remarks']).toEqual([]);
  });

  it('records no unknown labels for a standard report', () => {
    expect(parsed.unknownLabels).toEqual([]);
  });
});

describe('parseReport — inline values', () => {
  it('reads an inline value written with a space after the colon', () => {
    const parsed = parseReport(fixture('[ENGINEERING] AMIR-HP_.txt'));
    expect(parsed.fields['Antivirus status']).toEqual(['NORTON ACTIVE']);
  });

  it('reads an inline value written with no space after the colon', () => {
    const parsed = parseReport(fixture('[QAQC FAIRUS]HPFL05_.txt'));
    expect(parsed.fields['Antivirus status']).toEqual(['NORTON INSTALLED (ACTIVE)']);
  });
});

describe('parseReport — CRLF and blank reports', () => {
  it('handles CRLF line endings without leaving stray returns', () => {
    const parsed = parseReport(fixture('CARMEN-HP_.txt'));
    expect(parsed.fields['Antivirus status']).toEqual(['NORTON INSTALLED (7 DAYS)']);
  });

  it('parses a report whose every field is empty', () => {
    const parsed = parseReport(fixture('CARMEN-HP_.txt'));
    expect(parsed.isReport).toBe(true);
    expect(parsed.fields['Computer Name']).toEqual([]);
    expect(parsed.fields['Processor']).toEqual([]);
  });
});

describe('parseReport — not a report', () => {
  it('flags a file with no known label', () => {
    const parsed = parseReport('Dear team,\n\nPlease find the invoice attached.\n');
    expect(parsed.isReport).toBe(false);
  });
});

describe('parseReport — unknown labels', () => {
  it('records a label the scan script did not used to write', () => {
    const parsed = parseReport('Computer Name:\nPC1\n\nBitLocker Status:\nEnabled\n');
    expect(parsed.unknownLabels).toEqual([{ label: 'BitLocker Status', value: 'Enabled' }]);
  });

  it('does not record a pipe-delimited value as an unknown label', () => {
    const parsed = parseReport('RAM Slot Info:\nTotal Slots: 2 | Used Slots: 2\n');
    expect(parsed.unknownLabels).toEqual([]);
  });

  it('does not record a drive letter as an unknown label', () => {
    const parsed = parseReport('Server folder:\nY: | \\\\server\\PMW\n');
    expect(parsed.unknownLabels).toEqual([]);
  });

  it('does not record a Windows path as an unknown label', () => {
    const parsed = parseReport('Remarks:\nC:\\Users\\User\\Desktop\n');
    expect(parsed.unknownLabels).toEqual([]);
  });
});

describe('parseReport — blank lines inside a block', () => {
  it('does not truncate a block at a blank line', () => {
    const parsed = parseReport('GPU:\nIntel HD\n\nNVIDIA RTX\n\nTotal RAM:\n8 GB\n');
    expect(parsed.fields['GPU']).toEqual(['Intel HD', 'NVIDIA RTX']);
    expect(parsed.fields['Total RAM']).toEqual(['8 GB']);
  });
});
```

- [ ] **Step 3: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/parse/parseReport.test.js`
Expected: FAIL — `Failed to resolve import "./parseReport.js"`

- [ ] **Step 4: Write `parseReport.js`**

```js
import { KNOWN_LABELS, matchLabel } from './labels.js';

/** Lines the scan writes as decoration. They are not values. */
const SEPARATOR = /^=+$/;
const BANNERS = new Set(['COMPUTER INFORMATION', 'END OF REPORT']);

/**
 * A line that looks like `Something: value` but is not a known label. Used to
 * surface fields a future version of the scan script might add.
 */
const LABEL_SHAPED = /^([A-Za-z][\w /&()+.'-]{0,60}):\s*(.*)$/;

/** `C:\Users\...` and `Y: | \\server\...` are values, not labels. */
const DRIVE_LETTER = /^[A-Za-z]:/;

const INVISIBLE = /[\u00a0\u200b\u200c]/g;

function normaliseText(text) {
  return text
    .replace(/^\ufeff/, '')
    .replace(/\r\n?/g, '\n')
    .replace(INVISIBLE, ' ');
}

export function parseReport(text) {
  const fields = Object.fromEntries(KNOWN_LABELS.map((label) => [label, []]));
  const unknownLabels = [];
  const warnings = [];

  let current = null;
  let sawKnownLabel = false;

  for (const raw of normaliseText(text).split('\n')) {
    const line = raw.replace(/\s+$/, '').trim();

    if (!line) continue;
    if (SEPARATOR.test(line) || BANNERS.has(line)) continue;

    const hit = matchLabel(line);
    if (hit) {
      sawKnownLabel = true;
      current = hit.label;
      if (hit.inline) fields[current].push(hit.inline);
      continue;
    }

    // Not a known label. Before treating it as a value, check whether it looks
    // like a label the scan script has newly started writing — but never treat
    // a pipe-delimited line or a drive path as one, because those are the two
    // shapes real values take in this format.
    const shaped = LABEL_SHAPED.exec(line);
    if (shaped && !line.includes(' | ') && !DRIVE_LETTER.test(line)) {
      unknownLabels.push({ label: shaped[1].trim(), value: shaped[2].trim() });
      continue;
    }

    if (current) fields[current].push(line);
    else warnings.push(`Value before any label: ${line}`);
  }

  return { fields, unknownLabels, warnings, isReport: sawKnownLabel };
}
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/parse/parseReport.test.js`
Expected: PASS — 15 tests.

- [ ] **Step 6: Run the whole suite**

Run: `npm test`
Expected: PASS — the 73 pre-existing tests plus the new ones.

- [ ] **Step 7: Commit**

```bash
git add src/features/devices/
git commit -m "Parse device reports off a known-label whitelist"
```

---

## Task 3: `parseValues` — the pipe-delimited sub-structure

Implements spec §7.6.

**Files:**
- Create: `src/features/devices/parse/parseValues.js`
- Test: `src/features/devices/parse/parseValues.test.js`

**Interfaces:**
- Consumes: `cleanValue` from Task 1.
- Produces:
  - `parseSize(text): number | null` — `"477 GB"` → `477`, `"1 TB"` → `1024`
  - `parseRamSlots(lines): { sticks: {sizeGB,type,speedMhz,vendor,partNumber}[], totalSlots: number|null, usedSlots: number|null }`
  - `parseDrives(lines): { model, type, sizeGB, mechanical }[]`
  - `parseNetwork(lines): { connection, ssid, ip, assignment } | null`
  - `parseAntivirus(lines): { product, enabled }[]`
  - `parsePairs(lines): { left, right }[]`
  - `parseOffice(lines): string[]`
  - `parseGpus(lines): string[]`
  - `parseMonitors(lines): string[]`
  - `parseMailFiles(lines): { file, path, kind }[]`

- [ ] **Step 1: Write the failing test**

`src/features/devices/parse/parseValues.test.js`:

```js
import { describe, it, expect } from 'vitest';
import {
  parseSize, parseRamSlots, parseDrives, parseNetwork, parseAntivirus,
  parsePairs, parseOffice, parseGpus, parseMonitors, parseMailFiles,
} from './parseValues.js';

describe('parseSize', () => {
  it('reads GB and TB', () => {
    expect(parseSize('477 GB')).toBe(477);
    expect(parseSize('8 GB')).toBe(8);
    expect(parseSize('1 TB')).toBe(1024);
  });
  it('returns null for anything unparseable', () => {
    expect(parseSize('')).toBe(null);
    expect(parseSize('Unknown')).toBe(null);
  });
});

describe('parseRamSlots', () => {
  it('reads two sticks and the summary line', () => {
    const result = parseRamSlots([
      '4 GB | DDR4 | 3200 MHz | Samsung | M471A5244CB0-CWE',
      '4 GB | DDR4 | 3200 MHz | Samsung | M471A5244CB0-CWE',
      'Total Slots: 2 | Used Slots: 2',
    ]);
    expect(result.sticks).toHaveLength(2);
    expect(result.sticks[0]).toEqual({
      sizeGB: 4, type: 'DDR4', speedMhz: 3200,
      vendor: 'Samsung', partNumber: 'M471A5244CB0-CWE',
    });
    expect(result.totalSlots).toBe(2);
    expect(result.usedSlots).toBe(2);
  });

  it('falls back to counting sticks when Used Slots is blank', () => {
    const result = parseRamSlots([
      '16 GB | DDR4 | 3200 MHz | Samsung | M471A2G43AB2-CWE',
      'Total Slots: 2 | Used Slots: ',
    ]);
    expect(result.totalSlots).toBe(2);
    expect(result.usedSlots).toBe(1);
  });

  it('nulls the unset SMBIOS vendor and part number', () => {
    const result = parseRamSlots([
      '2 GB | Unknown | 333 MHz | Manufacturer1 | PartNum1',
      'Total Slots: 2 | Used Slots: ',
    ]);
    expect(result.sticks[0].vendor).toBe(null);
    expect(result.sticks[0].partNumber).toBe(null);
    expect(result.sticks[0].type).toBe(null);
    expect(result.sticks[0].sizeGB).toBe(2);
  });

  it('never counts the summary line as a stick', () => {
    const result = parseRamSlots(['Total Slots: 4 | Used Slots: ']);
    expect(result.sticks).toEqual([]);
    expect(result.usedSlots).toBe(0);
  });
});

describe('parseDrives', () => {
  it('reads model, type and size', () => {
    expect(parseDrives(['KBG50ZNV512G KIOXIA | SSD | 477 GB'])).toEqual([
      { model: 'KBG50ZNV512G KIOXIA', type: 'SSD', sizeGB: 477, mechanical: false },
    ]);
  });

  it('treats Unspecified as a mechanical disk', () => {
    const [drive] = parseDrives(['WDC WD10 JPVX-60JC3T1 | Unspecified | 932 GB']);
    expect(drive.type).toBe('HDD (assumed)');
    expect(drive.mechanical).toBe(true);
  });
});

describe('parseNetwork', () => {
  it('strips the inner SSID and IP prefixes', () => {
    expect(parseNetwork(['Wi-Fi | SSID: PMW_Group 7 | IP: 192.168.1.170 | Dynamic'])).toEqual({
      connection: 'Wi-Fi', ssid: 'PMW_Group 7', ip: '192.168.1.170', assignment: 'Dynamic',
    });
  });
  it('returns null when the block is empty', () => {
    expect(parseNetwork([])).toBe(null);
  });
});

describe('parseAntivirus', () => {
  it('de-duplicates repeated products and keeps enabled if any entry is enabled', () => {
    const result = parseAntivirus([
      'HP Wolf Pro Security | Enabled',
      'HP Wolf Pro Security | Disabled',
      'HP Wolf Pro Security | Enabled',
      'Norton 360 | Enabled',
      'Windows Defender | Disabled',
    ]);
    expect(result).toEqual([
      { product: 'HP Wolf Pro Security', enabled: true },
      { product: 'Norton 360', enabled: true },
      { product: 'Windows Defender', enabled: false },
    ]);
  });
});

describe('parsePairs', () => {
  it('splits a two-part line', () => {
    expect(parsePairs(['HP | 8BB6'])).toEqual([{ left: 'HP', right: '8BB6' }]);
  });
  it('keeps the whole line as left when there is no pipe', () => {
    expect(parsePairs(['server'])).toEqual([{ left: 'server', right: null }]);
  });
});

describe('parseOffice', () => {
  it('splits the single comma-separated line', () => {
    expect(parseOffice(['O365BusinessRetail,O365HomePremRetail']))
      .toEqual(['O365BusinessRetail', 'O365HomePremRetail']);
  });
});

describe('parseGpus', () => {
  it('drops the AnyDesk virtual display', () => {
    expect(parseGpus(['Intel(R) Iris(R) Xe Graphics', 'VirtualMonitorDriver Device']))
      .toEqual(['Intel(R) Iris(R) Xe Graphics']);
  });
});

describe('parseMonitors', () => {
  it('drops the Windows pseudo-monitor', () => {
    expect(parseMonitors(['Generic PnP Monitor', 'Default Monitor']))
      .toEqual(['Generic PnP Monitor']);
  });
});

describe('parseMailFiles', () => {
  it('classifies .ost as a mailbox and .pst as an archive', () => {
    const result = parseMailFiles([
      'ashraf@pmw-group.com.ost | C:\\Users\\User\\AppData\\Local\\Microsoft\\Outlook\\a.ost',
      'ashraf@pmw-industries.com.pst | C:\\Users\\User\\Documents\\Outlook Files\\b.pst',
    ]);
    expect(result.map((r) => r.kind)).toEqual(['mailbox', 'archive']);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/parse/parseValues.test.js`
Expected: FAIL — `Failed to resolve import "./parseValues.js"`

- [ ] **Step 3: Write `parseValues.js`**

```js
import { cleanValue } from './placeholders.js';

const split = (line) => line.split('|').map((part) => part.trim());

/** `Total Slots: 2 | Used Slots: 2` — a summary, not a stick. */
const SLOT_SUMMARY = /^total\s+slots\s*:/i;

export function parseSize(text) {
  const match = /(\d+(?:\.\d+)?)\s*(TB|GB|MB)/i.exec(text ?? '');
  if (!match) return null;
  const value = Number(match[1]);
  const unit = match[2].toUpperCase();
  if (unit === 'TB') return Math.round(value * 1024);
  if (unit === 'MB') return Math.round(value / 1024);
  return Math.round(value);
}

const parseInteger = (text) => {
  const match = /(\d+)/.exec(text ?? '');
  return match ? Number(match[1]) : null;
};

export function parseRamSlots(lines) {
  const sticks = [];
  let totalSlots = null;
  let usedSlots = null;

  for (const line of lines) {
    if (SLOT_SUMMARY.test(line)) {
      const [totalPart, usedPart = ''] = split(line);
      totalSlots = parseInteger(totalPart);
      usedSlots = parseInteger(usedPart);
      continue;
    }
    const [size, type, speed, vendor, partNumber] = split(line);
    sticks.push({
      sizeGB: parseSize(size),
      type: cleanValue(type),
      speedMhz: parseInteger(speed),
      vendor: cleanValue(vendor),
      partNumber: cleanValue(partNumber),
    });
  }

  // The scan leaves `Used Slots:` blank on 5 of 17 machines. Counting the
  // sticks it did report is strictly better than reporting nothing.
  if (usedSlots === null) usedSlots = sticks.length;

  return { sticks, totalSlots, usedSlots };
}

export function parseDrives(lines) {
  return lines.map((line) => {
    const [model, type, size] = split(line);
    // "Unspecified" means Win32_DiskDrive could not read MediaType. On every
    // machine in the sample set that is a spinning disk.
    const isSsd = /ssd/i.test(type ?? '');
    return {
      model: cleanValue(model),
      type: isSsd ? 'SSD' : 'HDD (assumed)',
      sizeGB: parseSize(size),
      mechanical: !isSsd,
    };
  });
}

const stripPrefix = (text, prefix) =>
  cleanValue((text ?? '').replace(new RegExp(`^${prefix}\\s*:\\s*`, 'i'), ''));

export function parseNetwork(lines) {
  if (!lines.length) return null;
  const [connection, ssid, ip, assignment] = split(lines[0]);
  return {
    connection: cleanValue(connection),
    ssid: stripPrefix(ssid, 'SSID'),
    ip: stripPrefix(ip, 'IP'),
    assignment: cleanValue(assignment),
  };
}

export function parseAntivirus(lines) {
  const byProduct = new Map();
  for (const line of lines) {
    const [product, state] = split(line);
    const name = cleanValue(product);
    if (!name) continue;
    const enabled = /enabled/i.test(state ?? '');
    // AMIR-HP lists HP Wolf Pro Security 22 times with conflicting states.
    // A product is protecting the machine if any of its entries is enabled.
    byProduct.set(name, (byProduct.get(name) ?? false) || enabled);
  }
  return [...byProduct].map(([product, enabled]) => ({ product, enabled }));
}

export function parsePairs(lines) {
  return lines.map((line) => {
    const [left, right = null] = split(line);
    return { left: cleanValue(left), right: right === null ? null : cleanValue(right) };
  });
}

export function parseOffice(lines) {
  return lines
    .flatMap((line) => line.split(','))
    .map(cleanValue)
    .filter(Boolean);
}

const REJECT_GPU = /virtualmonitordriver/i;
const REJECT_MONITOR = /^default monitor$/i;

export function parseGpus(lines) {
  return lines.map(cleanValue).filter((v) => v && !REJECT_GPU.test(v));
}

export function parseMonitors(lines) {
  return lines.map(cleanValue).filter((v) => v && !REJECT_MONITOR.test(v));
}

export function parseMailFiles(lines) {
  return parsePairs(lines).map(({ left, right }) => ({
    file: left,
    path: right,
    kind: /\.pst$/i.test(left ?? '') ? 'archive' : 'mailbox',
  }));
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/parse/parseValues.test.js`
Expected: PASS — 17 tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/parse/
git commit -m "Parse the pipe-delimited sub-structure of device report answers"
```

---

# PHASE 2 — Derivation

---

## Task 4: `deriveRam`

Implements spec §8.1.

**Files:**
- Create: `src/features/devices/derive/deriveRam.js`
- Test: `src/features/devices/derive/deriveRam.test.js`

**Interfaces:**
- Consumes: `parseRamSlots`, `parseSize` from Task 3.
- Produces: `deriveRam(slotLines: string[], totalRamLines: string[]): { installedRamGB, reportedRamGB, ramDiscrepancy, ramType, ramSpeedMhz, ramSlotsUsed, ramSlotsTotal, ramUpgradable }`

- [ ] **Step 1: Write the failing test**

`src/features/devices/derive/deriveRam.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { deriveRam } from './deriveRam.js';

const twoByFour = [
  '4 GB | DDR4 | 3200 MHz | Samsung | M471A5244CB0-CWE',
  '4 GB | DDR4 | 3200 MHz | Samsung | M471A5244CB0-CWE',
  'Total Slots: 2 | Used Slots: 2',
];

describe('deriveRam', () => {
  it('sums the sticks for the installed figure', () => {
    expect(deriveRam(twoByFour, ['8 GB']).installedRamGB).toBe(8);
  });

  it('flags the iGPU reservation gap: 15 GB reported is a 16 GB machine', () => {
    const result = deriveRam(
      ['8 GB | DDR4 | 3200 MHz | Micron Technology | 4ATF1G64HZ-3G2F1',
       '8 GB | DDR4 | 3200 MHz | Micron Technology | 4ATF1G64HZ-3G2F1',
       'Total Slots: 2 | Used Slots: 2'],
      ['15 GB'],
    );
    expect(result.installedRamGB).toBe(16);
    expect(result.reportedRamGB).toBe(15);
    expect(result.ramDiscrepancy).toBe(true);
  });

  it('does not flag a discrepancy when the two agree', () => {
    expect(deriveRam(twoByFour, ['8 GB']).ramDiscrepancy).toBe(false);
  });

  it('reports a free slot as upgradable', () => {
    const result = deriveRam(
      ['8 GB | DDR4 | 3200 MHz | Kingston | HP32D4S2S8MR-8', 'Total Slots: 2 | Used Slots: '],
      ['8 GB'],
    );
    expect(result.ramSlotsUsed).toBe(1);
    expect(result.ramSlotsTotal).toBe(2);
    expect(result.ramUpgradable).toBe(true);
  });

  it('is not upgradable when every slot is filled', () => {
    expect(deriveRam(twoByFour, ['8 GB']).ramUpgradable).toBe(false);
  });

  it('takes the slowest speed when sticks differ', () => {
    const result = deriveRam(
      ['8 GB | DDR4 | 3200 MHz | A | 1', '8 GB | DDR4 | 2667 MHz | B | 2'],
      ['16 GB'],
    );
    expect(result.ramSpeedMhz).toBe(2667);
  });

  it('reports the most common stick type', () => {
    expect(deriveRam(twoByFour, ['8 GB']).ramType).toBe('DDR4');
  });

  it('returns Unknown type when the scan could not read it', () => {
    const result = deriveRam(
      ['2 GB | Unknown | 333 MHz | Manufacturer1 | PartNum1', 'Total Slots: 2 | Used Slots: '],
      ['2 GB'],
    );
    expect(result.ramType).toBe('Unknown');
    expect(result.installedRamGB).toBe(2);
  });

  it('handles a report with no RAM block at all', () => {
    const result = deriveRam([], []);
    expect(result.installedRamGB).toBe(null);
    expect(result.reportedRamGB).toBe(null);
    expect(result.ramDiscrepancy).toBe(false);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/derive/deriveRam.test.js`
Expected: FAIL — `Failed to resolve import "./deriveRam.js"`

- [ ] **Step 3: Write `deriveRam.js`**

```js
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

  const sizes = sticks.map((s) => s.sizeGB).filter((n) => typeof n === 'number');
  const installedRamGB = sizes.length ? sizes.reduce((a, b) => a + b, 0) : null;
  const reportedRamGB = totalRamLines.length ? parseSize(totalRamLines[0]) : null;

  const speeds = sticks.map((s) => s.speedMhz).filter((n) => typeof n === 'number');
  const types = sticks.map((s) => s.type).filter(Boolean);

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
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/derive/deriveRam.test.js`
Expected: PASS — 9 tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/derive/
git commit -m "Derive installed RAM from slots rather than the reported total"
```

---

## Task 5: `deriveStorage`

Implements spec §8.2.

**Files:**
- Create: `src/features/devices/derive/deriveStorage.js`
- Test: `src/features/devices/derive/deriveStorage.test.js`

**Interfaces:**
- Consumes: `parseDrives` from Task 3.
- Produces: `deriveStorage(driveLines: string[]): { storageTotalGB, driveCount, hasHdd, storageType }`
  - `storageType` is one of `'SSD only' | 'Mixed' | 'HDD only' | 'Unknown'`.

- [ ] **Step 1: Write the failing test**

`src/features/devices/derive/deriveStorage.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { deriveStorage } from './deriveStorage.js';

describe('deriveStorage', () => {
  it('classifies a single SSD machine', () => {
    expect(deriveStorage(['KBG50ZNV512G KIOXIA | SSD | 477 GB'])).toEqual({
      storageTotalGB: 477, driveCount: 1, hasHdd: false, storageType: 'SSD only',
    });
  });

  it('sums both drives and calls an SSD plus a spinning disk Mixed', () => {
    const result = deriveStorage([
      'WDC WD10 JPVX-60JC3T1 | Unspecified | 932 GB',
      'SAMSUNG MZVLQ512HBLU-00BH1 | SSD | 477 GB',
    ]);
    expect(result.storageTotalGB).toBe(1409);
    expect(result.driveCount).toBe(2);
    expect(result.hasHdd).toBe(true);
    expect(result.storageType).toBe('Mixed');
  });

  it('calls a machine with only spinning disks HDD only', () => {
    expect(deriveStorage(['WDC WD10 | Unspecified | 932 GB']).storageType).toBe('HDD only');
  });

  it('returns Unknown for a report with no storage block', () => {
    expect(deriveStorage([])).toEqual({
      storageTotalGB: null, driveCount: 0, hasHdd: false, storageType: 'Unknown',
    });
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/derive/deriveStorage.test.js`
Expected: FAIL — `Failed to resolve import "./deriveStorage.js"`

- [ ] **Step 3: Write `deriveStorage.js`**

```js
import { parseDrives } from '../parse/parseValues.js';

export function deriveStorage(driveLines) {
  const drives = parseDrives(driveLines);

  if (!drives.length) {
    return { storageTotalGB: null, driveCount: 0, hasHdd: false, storageType: 'Unknown' };
  }

  const sizes = drives.map((d) => d.sizeGB).filter((n) => typeof n === 'number');
  const mechanical = drives.filter((d) => d.mechanical).length;

  let storageType = 'SSD only';
  if (mechanical === drives.length) storageType = 'HDD only';
  else if (mechanical > 0) storageType = 'Mixed';

  return {
    storageTotalGB: sizes.length ? sizes.reduce((a, b) => a + b, 0) : null,
    driveCount: drives.length,
    hasHdd: mechanical > 0,
    storageType,
  };
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/derive/deriveStorage.test.js`
Expected: PASS — 4 tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/derive/
git commit -m "Classify device storage as SSD only, mixed or mechanical"
```

---

## Task 6: `deriveCpu`

Implements spec §8.3 (the CPU half).

**Files:**
- Create: `src/features/devices/derive/deriveCpu.js`
- Test: `src/features/devices/derive/deriveCpu.test.js`

**Interfaces:**
- Consumes: `cleanValue` from Task 1.
- Produces: `deriveCpu(processorLines: string[], ramType: string|null): { cpuModel, cpuVendor, cpuGeneration, cpuAgeBand }`
  - `cpuVendor` is `'Intel' | 'AMD' | 'Other' | null`.
  - `cpuGeneration` is a display string: `'13'`, `'Ultra 1'`, `'Ryzen 7000'`, or `null`.
  - `cpuAgeBand` is `'Current' | 'Aging' | 'Obsolete' | 'Unknown'`.

**Why the SKU rule is not "take the first digit":** Intel's 4-digit SKUs are
ambiguous. `i7-3667U` is 3rd gen (first digit) but `i7-1355U` is 13th gen (first
two). A 4-digit number beginning `10`–`14` is a 10th-generation-or-later part;
anything beginning `2`–`9` is that generation. The `"13th Gen"` prefix the scan
usually writes is checked first because it needs no inference at all.

- [ ] **Step 1: Write the failing test**

`src/features/devices/derive/deriveCpu.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { deriveCpu } from './deriveCpu.js';

describe('deriveCpu — generation', () => {
  it('uses the explicit "Nth Gen" prefix when the scan writes one', () => {
    expect(deriveCpu(['13th Gen Intel(R) Core(TM) i7-1355U'], 'DDR4').cpuGeneration).toBe('13');
    expect(deriveCpu(['12th Gen Intel(R) Core(TM) i5-12400'], 'DDR4').cpuGeneration).toBe('12');
  });

  it('reads a 4-digit SKU beginning 10-14 as a two-digit generation', () => {
    expect(deriveCpu(['Intel(R) Core(TM) i5-1035G1 CPU @ 1.00GHz'], 'DDR4').cpuGeneration)
      .toBe('10');
  });

  it('reads a 4-digit SKU beginning 2-9 as a one-digit generation', () => {
    expect(deriveCpu(['Intel(R) Core(TM) i7-3667U CPU @ 2.00GHz'], 'DDR3').cpuGeneration)
      .toBe('3');
  });

  it('reads Core Ultra as a series, not an i-series generation', () => {
    expect(deriveCpu(['Intel(R) Core(TM) Ultra 5 125U'], 'DDR5').cpuGeneration).toBe('Ultra 1');
  });

  it('reads an AMD Ryzen series', () => {
    expect(deriveCpu(['AMD Ryzen 5 7430U with Radeon Graphics   '], 'DDR4').cpuGeneration)
      .toBe('Ryzen 7000');
  });

  it('has no generation for a Pentium', () => {
    expect(deriveCpu(['Intel(R) Pentium(R) Dual  CPU  E2160  @ 1.80GHz'], null).cpuGeneration)
      .toBe(null);
  });
});

describe('deriveCpu — vendor and model', () => {
  it('reads the vendor', () => {
    expect(deriveCpu(['13th Gen Intel(R) Core(TM) i7-1355U'], 'DDR4').cpuVendor).toBe('Intel');
    expect(deriveCpu(['AMD Ryzen 5 7430U'], 'DDR4').cpuVendor).toBe('AMD');
  });

  it('trims the trailing whitespace the scan writes after AMD names', () => {
    expect(deriveCpu(['AMD Ryzen 5 7430U with Radeon Graphics         '], 'DDR4').cpuModel)
      .toBe('AMD Ryzen 5 7430U with Radeon Graphics');
  });
});

describe('deriveCpu — age band', () => {
  it('calls 10th generation and later Current', () => {
    expect(deriveCpu(['13th Gen Intel(R) Core(TM) i7-1355U'], 'DDR4').cpuAgeBand).toBe('Current');
    expect(deriveCpu(['Intel(R) Core(TM) i5-1035G1'], 'DDR4').cpuAgeBand).toBe('Current');
  });

  it('calls Core Ultra Current', () => {
    expect(deriveCpu(['Intel(R) Core(TM) Ultra 5 125U'], 'DDR5').cpuAgeBand).toBe('Current');
  });

  it('calls 7th to 9th generation Aging', () => {
    expect(deriveCpu(['Intel(R) Core(TM) i5-8250U'], 'DDR4').cpuAgeBand).toBe('Aging');
  });

  it('calls 6th generation and earlier Obsolete', () => {
    expect(deriveCpu(['Intel(R) Core(TM) i7-3667U CPU @ 2.00GHz'], 'DDR3').cpuAgeBand)
      .toBe('Obsolete');
  });

  it('calls a Pentium with no generation Obsolete', () => {
    expect(deriveCpu(['Intel(R) Pentium(R) Dual  CPU  E2160  @ 1.80GHz'], null).cpuAgeBand)
      .toBe('Obsolete');
  });

  it('ranks AMD by series rather than inventing an Intel-comparable generation', () => {
    expect(deriveCpu(['AMD Ryzen 5 7430U'], 'DDR4').cpuAgeBand).toBe('Current');
    expect(deriveCpu(['AMD Ryzen 5 3500U'], 'DDR4').cpuAgeBand).toBe('Aging');
    expect(deriveCpu(['AMD Ryzen 3 2200U'], 'DDR4').cpuAgeBand).toBe('Obsolete');
  });

  it('calls DDR3 Obsolete even when no generation could be read', () => {
    expect(deriveCpu(['Some Unknown CPU'], 'DDR3').cpuAgeBand).toBe('Obsolete');
  });

  it('returns Unknown when there is nothing to go on', () => {
    expect(deriveCpu([], null).cpuAgeBand).toBe('Unknown');
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/derive/deriveCpu.test.js`
Expected: FAIL — `Failed to resolve import "./deriveCpu.js"`

- [ ] **Step 3: Write `deriveCpu.js`**

```js
import { cleanValue } from '../parse/placeholders.js';

const OBSOLETE_FAMILIES = /pentium|celeron|atom/i;

/**
 * Intel 4-digit SKUs are ambiguous: i7-3667U is 3rd generation (first digit)
 * and i7-1355U is 13th (first two). A 4-digit number starting 10-14 is a
 * 10th-generation-or-later part; 2-9 is that generation.
 */
function intelGenerationFromSku(sku) {
  if (sku.length >= 5) return Number(sku.slice(0, 2));
  if (sku.length === 4) {
    const leading = Number(sku.slice(0, 2));
    return leading >= 10 && leading <= 14 ? leading : Number(sku[0]);
  }
  return null;
}

function readGeneration(model) {
  // The scan usually writes it outright — no inference needed.
  const explicit = /(\d{1,2})(?:st|nd|rd|th)\s+Gen/i.exec(model);
  if (explicit) return { kind: 'intel', value: Number(explicit[1]) };

  const ultra = /Core\(TM\)\s+Ultra\s+\d+\s+(\d)\d{2}/i.exec(model);
  if (ultra) return { kind: 'ultra', value: Number(ultra[1]) };

  const core = /i[3579][- ](\d{4,5})/i.exec(model);
  if (core) return { kind: 'intel', value: intelGenerationFromSku(core[1]) };

  const ryzen = /Ryzen\s+\d+\s+(\d)\d{3}/i.exec(model);
  if (ryzen) return { kind: 'amd', value: Number(ryzen[1]) };

  return { kind: 'none', value: null };
}

function readAgeBand(model, generation, ramType) {
  if (ramType && /^DDR[12]$|^DDR3$/i.test(ramType)) return 'Obsolete';

  if (generation.kind === 'ultra') return 'Current';

  if (generation.kind === 'intel' && generation.value) {
    if (generation.value >= 10) return 'Current';
    if (generation.value >= 7) return 'Aging';
    return 'Obsolete';
  }

  if (generation.kind === 'amd' && generation.value) {
    // AMD mobile numbering does not map onto Intel generations — a 7430U is a
    // Zen 3 part wearing a 7000 badge — so it is ranked on its own series.
    if (generation.value >= 5) return 'Current';
    if (generation.value >= 3) return 'Aging';
    return 'Obsolete';
  }

  if (OBSOLETE_FAMILIES.test(model)) return 'Obsolete';
  if (ramType && /^DDR5$/i.test(ramType)) return 'Current';
  if (ramType && /^DDR4$/i.test(ramType)) return 'Aging';
  return 'Unknown';
}

export function deriveCpu(processorLines, ramType) {
  const cpuModel = processorLines.length ? cleanValue(processorLines[0]) : null;

  if (!cpuModel) {
    return { cpuModel: null, cpuVendor: null, cpuGeneration: null, cpuAgeBand: 'Unknown' };
  }

  let cpuVendor = 'Other';
  if (/intel/i.test(cpuModel)) cpuVendor = 'Intel';
  else if (/amd|ryzen/i.test(cpuModel)) cpuVendor = 'AMD';

  const generation = readGeneration(cpuModel);

  let cpuGeneration = null;
  if (generation.kind === 'intel' && generation.value) cpuGeneration = String(generation.value);
  else if (generation.kind === 'ultra') cpuGeneration = `Ultra ${generation.value}`;
  else if (generation.kind === 'amd') cpuGeneration = `Ryzen ${generation.value}000`;

  return {
    cpuModel,
    cpuVendor,
    cpuGeneration,
    cpuAgeBand: readAgeBand(cpuModel, generation, ramType),
  };
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/derive/deriveCpu.test.js`
Expected: PASS — 16 tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/derive/
git commit -m "Read CPU generation from Intel, Core Ultra and Ryzen naming"
```

---

## Task 7: `deriveIdentity` — owner, department, device type

Implements spec §8.3 (device type) and §8.4.

**Files:**
- Create: `src/features/devices/derive/deriveIdentity.js`
- Test: `src/features/devices/derive/deriveIdentity.test.js`

**Interfaces:**
- Consumes: `cleanValue` (Task 1), `parsePairs`, `parseMailFiles` (Task 3).
- Produces:
  - `KNOWN_DEPARTMENTS: string[]`
  - `parseFileName(fileName): { bracket: string|null, stem: string }`
  - `deriveIdentity(fields, fileName): { computerName, owner, ownerSource, department, deviceType, deviceTypeConfident }`
  - `ownerSource` is `'Name field' | 'Filename' | 'Server credential' | 'Email' | null`.

- [ ] **Step 1: Write the failing test**

`src/features/devices/derive/deriveIdentity.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { deriveIdentity, parseFileName } from './deriveIdentity.js';

const emptyFields = {
  Name: [], 'Computer Name': [], 'Computer Model': [], Motherboard: [],
  'PMW Server and credentials': [],
  'Email data files found Active or Inactive account': [],
};
const withFields = (overrides) => ({ ...emptyFields, ...overrides });

describe('parseFileName', () => {
  it('splits a bracket followed by a space', () => {
    expect(parseFileName('[FINANCE] LEMON-HP_.txt'))
      .toEqual({ bracket: 'FINANCE', stem: 'LEMON-HP' });
  });

  it('splits a bracket with no space before the name', () => {
    expect(parseFileName('[QAQC FAIRUS]HPFL05_.txt'))
      .toEqual({ bracket: 'QAQC FAIRUS', stem: 'HPFL05' });
  });

  it('handles a filename with no bracket', () => {
    expect(parseFileName('ASHRAF-PC_.txt')).toEqual({ bracket: null, stem: 'ASHRAF-PC' });
  });
});

describe('deriveIdentity — department and owner from the bracket', () => {
  it('splits a bracket holding both a department and a person', () => {
    const result = deriveIdentity(withFields({ 'Computer Name': ['HPFL05'] }),
      '[QAQC FAIRUS]HPFL05_.txt');
    expect(result.department).toBe('QAQC');
    expect(result.owner).toBe('Fairus');
    expect(result.ownerSource).toBe('Filename');
  });

  it('keeps a two-word department whole', () => {
    const result = deriveIdentity(withFields({ 'Computer Name': ['PMWP001'] }),
      '[PML GUARDHOUSE] PMWP001_.txt');
    expect(result.department).toBe('PML GUARDHOUSE');
    expect(result.owner).toBe(null);
  });

  it('reads a department-only bracket', () => {
    const result = deriveIdentity(withFields({ 'Computer Name': ['AMIR-HP'] }),
      '[ENGINEERING] AMIR-HP_.txt');
    expect(result.department).toBe('ENGINEERING');
  });
});

describe('deriveIdentity — the owner chain', () => {
  it('prefers the Name field above everything', () => {
    const result = deriveIdentity(
      withFields({ Name: ['Siti Aminah'], 'PMW Server and credentials': ['server | ashraf'] }),
      '[SALES] X_.txt');
    expect(result.owner).toBe('Siti Aminah');
    expect(result.ownerSource).toBe('Name field');
  });

  it('falls back to the server credential username', () => {
    const result = deriveIdentity(
      withFields({ 'PMW Server and credentials': ['server | ashraf'] }), 'ASHRAF-PC_.txt');
    expect(result.owner).toBe('Ashraf');
    expect(result.ownerSource).toBe('Server credential');
  });

  it('falls back to the first mailbox local part, title-cased', () => {
    const result = deriveIdentity(
      withFields({
        'Email data files found Active or Inactive account': [
          'lemon.cheong@pmw-group.com.ost | C:\\Users\\user\\a.ost',
        ],
      }), 'LEMON-HP_.txt');
    expect(result.owner).toBe('Lemon Cheong');
    expect(result.ownerSource).toBe('Email');
  });

  it('prefers a mailbox over an archive', () => {
    const result = deriveIdentity(
      withFields({
        'Email data files found Active or Inactive account': [
          'old.account@pmw-industries.com.pst | C:\\a.pst',
          'jiva.ran@pmw-group.com.ost | C:\\b.ost',
        ],
      }), 'JIVA_.txt');
    expect(result.owner).toBe('Jiva Ran');
  });

  it('reports no owner when nothing in the chain yields one', () => {
    const result = deriveIdentity(emptyFields, 'PMWP001_.txt');
    expect(result.owner).toBe(null);
    expect(result.ownerSource).toBe(null);
  });
});

describe('deriveIdentity — computer name', () => {
  it('uses the field when present', () => {
    expect(deriveIdentity(withFields({ 'Computer Name': ['ASHRAF-PC'] }), 'x_.txt').computerName)
      .toBe('ASHRAF-PC');
  });

  it('falls back to the filename stem', () => {
    expect(deriveIdentity(emptyFields, '[FINANCE] EVONNE-HP_.txt').computerName)
      .toBe('EVONNE-HP');
  });
});

describe('deriveIdentity — device type', () => {
  it('reads a desktop board before anything else', () => {
    const result = deriveIdentity(withFields({
      'Computer Model': ['MS-7D99'],
      Motherboard: ['Micro-Star International Co., Ltd. | PRO B760M-A WIFI (MS-7D99)'],
    }), 'UMAIRAH-PC_.txt');
    expect(result.deviceType).toBe('Desktop');
  });

  it('does not trust a DESKTOP- computer name over the model', () => {
    const result = deriveIdentity(withFields({
      'Computer Name': ['DESKTOP-2A3ERS8'],
      'Computer Model': ['HP EliteBook Folio 9470m'],
      Motherboard: ['Hewlett-Packard | 18DF'],
    }), 'DESKTOP-2A3ERS8_.txt');
    expect(result.deviceType).toBe('Laptop');
  });

  it('reads a Dell Precision as a laptop', () => {
    const result = deriveIdentity(withFields({
      'Computer Model': ['Precision 3490'], Motherboard: ['Dell Inc. | 0JTMW8'],
    }), 'PMWL034_.txt');
    expect(result.deviceType).toBe('Laptop');
  });

  it('reads the unset DMI product string plus an ASUS board as a desktop', () => {
    const result = deriveIdentity(withFields({
      'Computer Model': ['System Product Name'],
      Motherboard: ['ASUSTeK COMPUTER INC. | PRIME H610M-K D4'],
    }), 'PMWP001_.txt');
    expect(result.deviceType).toBe('Desktop');
  });

  it('reports Unknown rather than guessing when there is no signal', () => {
    const result = deriveIdentity(emptyFields, 'CARMEN-HP_.txt');
    expect(result.deviceType).toBe('Unknown');
    expect(result.deviceTypeConfident).toBe(false);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/derive/deriveIdentity.test.js`
Expected: FAIL — `Failed to resolve import "./deriveIdentity.js"`

- [ ] **Step 3: Write `deriveIdentity.js`**

```js
import { cleanValue } from '../parse/placeholders.js';
import { parsePairs, parseMailFiles } from '../parse/parseValues.js';

/** Longest first, so `PML GUARDHOUSE` matches before `PML` ever could. */
export const KNOWN_DEPARTMENTS = [
  'PML GUARDHOUSE', 'STOCKYARDF1', 'ENGINEERING', 'PRODUCTION', 'PURCHASING',
  'MARKETING', 'SHIPPING', 'FINANCE', 'ACCOUNT', 'SALES', 'ADMIN', 'STORE',
  'QAQC', 'QC', 'HR', 'IT',
];

// No trailing \b: `MS-7D99` has no word boundary between the 7 and the D, so
// `\bMS-7\b` would fail on the exact string this rule exists to catch.
const DESKTOP_BOARD = /(PRIME|MS-7\d|P5G|PRO B\d|TUF|ROG|\bH\d{3}M|\bB\d{3}M)/i;
const LAPTOP_MODEL =
  /Laptop|Notebook|Book|Pavilion|Inspiron|Latitude|Vostro|ThinkPad|IdeaPad|Precision\s+\d{4}|Folio|Elite/i;

const titleCase = (text) =>
  text
    .replace(/[._-]+/g, ' ')
    .trim()
    .split(/\s+/)
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(' ');

export function parseFileName(fileName) {
  const withoutExtension = fileName.replace(/\.txt$/i, '');
  const bracketMatch = /^\s*\[([^\]]+)\]\s*/.exec(withoutExtension);

  const bracket = bracketMatch ? bracketMatch[1].trim() : null;
  const rest = bracketMatch ? withoutExtension.slice(bracketMatch[0].length) : withoutExtension;

  return { bracket, stem: rest.replace(/_+$/, '').trim() };
}

function splitBracket(bracket) {
  if (!bracket) return { department: null, person: null };

  const upper = bracket.toUpperCase();
  const department = KNOWN_DEPARTMENTS.find(
    (dept) => upper === dept || upper.startsWith(`${dept} `),
  );

  if (!department) return { department: bracket, person: null };

  const remainder = bracket.slice(department.length).trim();
  return { department, person: remainder ? titleCase(remainder) : null };
}

function resolveOwner(fields, person) {
  const named = fields['Name']?.length ? cleanValue(fields['Name'][0]) : null;
  if (named) return { owner: named, ownerSource: 'Name field' };

  if (person) return { owner: person, ownerSource: 'Filename' };

  const credentials = parsePairs(fields['PMW Server and credentials'] ?? []);
  const username = credentials.find((pair) => pair.right)?.right;
  if (username) return { owner: titleCase(username), ownerSource: 'Server credential' };

  const mail = parseMailFiles(fields['Email data files found Active or Inactive account'] ?? []);
  // An .ost is the signed-in mailbox; a .pst is an archive that may belong to
  // somebody who left, so it is a weaker signal for "who uses this machine".
  const primary = mail.find((entry) => entry.kind === 'mailbox') ?? mail[0];
  if (primary?.file) {
    const localPart = primary.file.split('@')[0];
    if (localPart) return { owner: titleCase(localPart), ownerSource: 'Email' };
  }

  return { owner: null, ownerSource: null };
}

function resolveDeviceType(fields) {
  const model = fields['Computer Model']?.length ? fields['Computer Model'][0] : '';
  const board = fields['Motherboard']?.length ? fields['Motherboard'][0] : '';

  // The board is checked first because the computer NAME lies:
  // DESKTOP-2A3ERS8 is an HP EliteBook laptop.
  if (DESKTOP_BOARD.test(board)) return { deviceType: 'Desktop', deviceTypeConfident: true };
  if (LAPTOP_MODEL.test(model)) return { deviceType: 'Laptop', deviceTypeConfident: true };

  // An unset DMI product string means nobody flashed a model name in, which in
  // practice means a desktop assembled from parts.
  if (/^system product name$/i.test(model.trim())) {
    return { deviceType: 'Desktop', deviceTypeConfident: false };
  }

  return { deviceType: 'Unknown', deviceTypeConfident: false };
}

export function deriveIdentity(fields, fileName) {
  const { bracket, stem } = parseFileName(fileName);
  const { department, person } = splitBracket(bracket);

  const fromField = fields['Computer Name']?.length ? cleanValue(fields['Computer Name'][0]) : null;

  return {
    computerName: fromField ?? (stem || null),
    department,
    ...resolveOwner(fields, person),
    ...resolveDeviceType(fields),
  };
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/derive/deriveIdentity.test.js`
Expected: PASS — 17 tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/derive/
git commit -m "Resolve device owner, department and type from four sources"
```

---

## Task 8: `deriveHealth` — OS support, antivirus, scan completeness

Implements spec §8.5, §8.6.

**Files:**
- Create: `src/features/devices/derive/deriveHealth.js`
- Test: `src/features/devices/derive/deriveHealth.test.js`

**Interfaces:**
- Consumes: `cleanValue` (Task 1), `parseAntivirus` (Task 3).
- Produces: `deriveHealth(fields): { windowsVersion, windowsMajor, windowsEdition, osSupported, antivirusStatus, antivirusStatusRaw, antivirusProducts, avProtected, scanComplete }`
  - `antivirusStatus` is `'Active' | 'Installed — Inactive' | 'Trial' | 'Not Installed' | 'Unknown'`. Note the em dash.

**Ordering trap:** `NORTON INSTALLED (DEACTIVATED)` contains the substring
`ACTIVAT`. The deactivated test must run before the active test or every
deactivated machine reads as protected.

- [ ] **Step 1: Write the failing test**

`src/features/devices/derive/deriveHealth.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { deriveHealth } from './deriveHealth.js';

const base = {
  'Computer Name': ['PC1'], Processor: ['Intel'], 'Storage Drives': ['D | SSD | 477 GB'],
  'Windows Version': [], 'Antivirus status': [], Antivirus: [],
};
const withFields = (overrides) => ({ ...base, ...overrides });

describe('deriveHealth — Windows', () => {
  it('reads major version and edition', () => {
    const result = deriveHealth(withFields({ 'Windows Version': ['Microsoft Windows 11 Pro'] }));
    expect(result.windowsMajor).toBe(11);
    expect(result.windowsEdition).toBe('Pro');
    expect(result.osSupported).toBe(true);
  });

  it('reads the long Home edition name', () => {
    const result = deriveHealth(withFields({
      'Windows Version': ['Microsoft Windows 11 Home Single Language'],
    }));
    expect(result.windowsEdition).toBe('Home Single Language');
  });

  it('marks Windows 10 unsupported — it lost support on 14 October 2025', () => {
    const result = deriveHealth(withFields({ 'Windows Version': ['Microsoft Windows 10 Pro'] }));
    expect(result.windowsMajor).toBe(10);
    expect(result.osSupported).toBe(false);
  });

  it('returns null support for a report with no Windows line', () => {
    expect(deriveHealth(withFields({})).osSupported).toBe(null);
  });
});

describe('deriveHealth — antivirus status', () => {
  const status = (raw, products = []) =>
    deriveHealth(withFields({ 'Antivirus status': raw ? [raw] : [], Antivirus: products }))
      .antivirusStatus;

  it('normalises every spelling the scan produces', () => {
    expect(status('NORTON NOT INSTALLED')).toBe('Not Installed');
    expect(status('NORTON ACTIVATED')).toBe('Active');
    expect(status('NORTON ACTIVE')).toBe('Active');
    expect(status('NORTON INSTALLED (ACTIVE)')).toBe('Active');
    expect(status('NORTON INSTALLED (7 DAYS)')).toBe('Trial');
  });

  it('does not read DEACTIVATED as active', () => {
    expect(status('NORTON INSTALLED (DEACTIVATED)')).toBe('Installed — Inactive');
  });

  it('falls back to the antivirus block when the status line is blank', () => {
    expect(status('', ['Norton 360 | Enabled'])).toBe('Active');
    expect(status('', ['Norton 360 | Disabled'])).toBe('Installed — Inactive');
    expect(status('', [])).toBe('Unknown');
  });
});

describe('deriveHealth — protection', () => {
  it('counts Windows Defender as protection', () => {
    const result = deriveHealth(withFields({
      'Antivirus status': ['NORTON NOT INSTALLED'],
      Antivirus: ['Windows Defender | Enabled'],
    }));
    expect(result.avProtected).toBe(true);
  });

  it('reports unprotected when every product is disabled', () => {
    const result = deriveHealth(withFields({
      Antivirus: ['Norton 360 | Disabled', 'Windows Defender | Disabled'],
    }));
    expect(result.avProtected).toBe(false);
  });

  it('de-duplicates the repeated products before judging', () => {
    const result = deriveHealth(withFields({
      Antivirus: Array(22).fill('HP Wolf Pro Security | Enabled'),
    }));
    expect(result.antivirusProducts).toEqual([{ product: 'HP Wolf Pro Security', enabled: true }]);
  });
});

describe('deriveHealth — scan completeness', () => {
  it('is complete when the core fields are present', () => {
    expect(deriveHealth(base).scanComplete).toBe(true);
  });

  it('is incomplete when name, processor and storage are all empty', () => {
    const result = deriveHealth({
      ...base, 'Computer Name': [], Processor: [], 'Storage Drives': [],
    });
    expect(result.scanComplete).toBe(false);
  });

  it('is complete when only one core field is missing', () => {
    expect(deriveHealth(withFields({ Processor: [] })).scanComplete).toBe(true);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/derive/deriveHealth.test.js`
Expected: FAIL — `Failed to resolve import "./deriveHealth.js"`

- [ ] **Step 3: Write `deriveHealth.js`**

```js
import { cleanValue } from '../parse/placeholders.js';
import { parseAntivirus } from '../parse/parseValues.js';

const INACTIVE = 'Installed — Inactive';

function readWindows(lines) {
  const windowsVersion = lines.length ? cleanValue(lines[0]) : null;
  if (!windowsVersion) {
    return { windowsVersion: null, windowsMajor: null, windowsEdition: null, osSupported: null };
  }

  const match = /Windows\s+(\d+)\s*(.*)$/i.exec(windowsVersion);
  const windowsMajor = match ? Number(match[1]) : null;

  return {
    windowsVersion,
    windowsMajor,
    windowsEdition: match && match[2] ? match[2].trim() : null,
    // Windows 10 reached end of support on 14 October 2025.
    osSupported: windowsMajor === null ? null : windowsMajor >= 11,
  };
}

function readAntivirusStatus(raw, products) {
  if (raw) {
    // `DEACTIVATED` contains `ACTIVAT`, so it has to be tested before `activ`.
    if (/not\s*installed/i.test(raw)) return 'Not Installed';
    if (/deactivat|disabled|expired/i.test(raw)) return INACTIVE;
    if (/\d+\s*days?|trial/i.test(raw)) return 'Trial';
    if (/activ|enabled/i.test(raw)) return 'Active';
  }

  if (!products.length) return 'Unknown';
  return products.some((entry) => entry.enabled) ? 'Active' : INACTIVE;
}

export function deriveHealth(fields) {
  const antivirusProducts = parseAntivirus(fields['Antivirus'] ?? []);
  const antivirusStatusRaw = fields['Antivirus status']?.length
    ? cleanValue(fields['Antivirus status'][0])
    : null;

  // A scan that failed early writes the header and nothing else. Importing it
  // as a machine with no CPU and no disk would drag down every fleet average,
  // so it is marked instead and excluded from the statistics.
  const scanComplete = !(
    !fields['Computer Name']?.length &&
    !fields['Processor']?.length &&
    !fields['Storage Drives']?.length
  );

  return {
    ...readWindows(fields['Windows Version'] ?? []),
    antivirusStatus: readAntivirusStatus(antivirusStatusRaw, antivirusProducts),
    antivirusStatusRaw,
    antivirusProducts,
    avProtected: antivirusProducts.some((entry) => entry.enabled),
    scanComplete,
  };
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/derive/deriveHealth.test.js`
Expected: PASS — 13 tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/derive/
git commit -m "Normalise Windows support and antivirus status from scan text"
```

---

## Task 9: `riskScore`

Implements spec §8.7.

**Files:**
- Create: `src/features/devices/derive/riskScore.js`
- Test: `src/features/devices/derive/riskScore.test.js`

**Interfaces:**
- Consumes: the fields produced by Tasks 4–8.
- Produces: `riskScore(device): { riskScore: number|null, riskLevel: string, riskReasons: string[] }`
  - `device` needs `{ osSupported, antivirusStatus, avProtected, installedRamGB, cpuAgeBand, hasHdd, scanComplete }`.
  - `riskLevel` is `'Critical' | 'High' | 'Watch' | 'OK' | 'Unknown'`.

- [ ] **Step 1: Write the failing test**

`src/features/devices/derive/riskScore.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { riskScore } from './riskScore.js';

const healthy = {
  osSupported: true, antivirusStatus: 'Active', avProtected: true,
  installedRamGB: 16, cpuAgeBand: 'Current', hasHdd: false, scanComplete: true,
};
const device = (overrides) => riskScore({ ...healthy, ...overrides });

describe('riskScore — individual signals', () => {
  it('scores a healthy machine zero', () => {
    expect(device({}).riskScore).toBe(0);
    expect(device({}).riskLevel).toBe('OK');
  });

  it('charges 40 for an unsupported OS', () => {
    expect(device({ osSupported: false }).riskScore).toBe(40);
  });

  it('charges 30 for missing antivirus', () => {
    expect(device({ antivirusStatus: 'Not Installed', avProtected: false }).riskScore).toBe(30);
  });

  it('charges 30 when a product is installed but nothing is enabled', () => {
    expect(device({ antivirusStatus: 'Installed — Inactive', avProtected: false }).riskScore)
      .toBe(30);
  });

  it('charges 15 for 8 GB and 25 for 4 GB, never both', () => {
    expect(device({ installedRamGB: 8 }).riskScore).toBe(15);
    expect(device({ installedRamGB: 4 }).riskScore).toBe(25);
    expect(device({ installedRamGB: 2 }).riskScore).toBe(25);
  });

  it('charges 25 for an obsolete CPU and 10 for an aging one', () => {
    expect(device({ cpuAgeBand: 'Obsolete' }).riskScore).toBe(25);
    expect(device({ cpuAgeBand: 'Aging' }).riskScore).toBe(10);
  });

  it('charges 10 for a mechanical disk', () => {
    expect(device({ hasHdd: true }).riskScore).toBe(10);
  });
});

describe('riskScore — bands', () => {
  it('places the boundaries at 20, 40 and 60', () => {
    expect(device({ hasHdd: true }).riskLevel).toBe('OK');            // 10
    expect(device({ cpuAgeBand: 'Obsolete' }).riskLevel).toBe('Watch'); // 25
    expect(device({ osSupported: false }).riskLevel).toBe('High');      // 40
    expect(device({ osSupported: false, cpuAgeBand: 'Obsolete' }).riskLevel)
      .toBe('Critical');                                                // 65
  });
});

describe('riskScore — the real machines', () => {
  it('scores DESKTOP-8SBR420 at 100 — Windows 10, Pentium, 2 GB, spinning disk', () => {
    const result = riskScore({
      osSupported: false, antivirusStatus: 'Active', avProtected: true,
      installedRamGB: 2, cpuAgeBand: 'Obsolete', hasHdd: true, scanComplete: true,
    });
    expect(result.riskScore).toBe(100);
    expect(result.riskLevel).toBe('Critical');
  });

  it('scores HPFL05 at 80 — Windows 10, 3rd gen DDR3, 8 GB', () => {
    const result = riskScore({
      osSupported: false, antivirusStatus: 'Active', avProtected: true,
      installedRamGB: 8, cpuAgeBand: 'Obsolete', hasHdd: false, scanComplete: true,
    });
    expect(result.riskScore).toBe(80);
  });

  it('scores AMIR-HP at 50 — Windows 10 plus a spinning disk', () => {
    const result = riskScore({
      osSupported: false, antivirusStatus: 'Active', avProtected: true,
      installedRamGB: 16, cpuAgeBand: 'Current', hasHdd: true, scanComplete: true,
    });
    expect(result.riskScore).toBe(50);
  });

  it('scores ASHRAF-PC at 15 — only its 8 GB counts against it', () => {
    expect(device({ installedRamGB: 8 }).riskScore).toBe(15);
  });
});

describe('riskScore — reasons and unknowns', () => {
  it('lists a reason for every charged signal', () => {
    const result = device({ osSupported: false, hasHdd: true });
    expect(result.riskReasons).toEqual([
      'Windows 10 or older — no security updates since 14 Oct 2025',
      'Mechanical hard disk',
    ]);
  });

  it('returns a null score for an incomplete scan rather than calling it healthy', () => {
    const result = riskScore({ ...healthy, scanComplete: false });
    expect(result.riskScore).toBe(null);
    expect(result.riskLevel).toBe('Unknown');
    expect(result.riskReasons).toEqual(['Scan incomplete — re-run the report']);
  });

  it('does not charge for a signal it cannot read', () => {
    const result = riskScore({
      osSupported: null, antivirusStatus: 'Unknown', avProtected: false,
      installedRamGB: null, cpuAgeBand: 'Unknown', hasHdd: false, scanComplete: true,
    });
    expect(result.riskScore).toBe(0);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/derive/riskScore.test.js`
Expected: FAIL — `Failed to resolve import "./riskScore.js"`

- [ ] **Step 3: Write `riskScore.js`**

```js
/**
 * Additive and explainable on purpose: the dashboard shows WHY a machine
 * scored what it did, so every charge records its reason.
 *
 * Signals the scan could not read charge nothing. An unknown antivirus state
 * is not evidence of an unprotected machine, and charging for it would push
 * every partially-readable report into the attention queue.
 */
const RULES = [
  {
    points: 40,
    reason: 'Windows 10 or older — no security updates since 14 Oct 2025',
    applies: (d) => d.osSupported === false,
  },
  {
    points: 30,
    reason: 'No active antivirus',
    applies: (d) => d.antivirusStatus !== 'Unknown' && !d.avProtected,
  },
  {
    points: 25,
    reason: '4 GB of RAM or less',
    applies: (d) => typeof d.installedRamGB === 'number' && d.installedRamGB <= 4,
  },
  {
    points: 15,
    reason: '8 GB of RAM or less',
    applies: (d) =>
      typeof d.installedRamGB === 'number' && d.installedRamGB > 4 && d.installedRamGB <= 8,
  },
  { points: 25, reason: 'Obsolete processor', applies: (d) => d.cpuAgeBand === 'Obsolete' },
  { points: 10, reason: 'Aging processor', applies: (d) => d.cpuAgeBand === 'Aging' },
  { points: 10, reason: 'Mechanical hard disk', applies: (d) => d.hasHdd === true },
];

function levelFor(score) {
  if (score >= 60) return 'Critical';
  if (score >= 40) return 'High';
  if (score >= 20) return 'Watch';
  return 'OK';
}

export function riskScore(device) {
  // An unscanned machine is unknown, not healthy. Scoring it zero would let a
  // failed scan sit at the top of the "all clear" list.
  if (device.scanComplete === false) {
    return {
      riskScore: null,
      riskLevel: 'Unknown',
      riskReasons: ['Scan incomplete — re-run the report'],
    };
  }

  const hits = RULES.filter((rule) => rule.applies(device));
  const score = hits.reduce((total, rule) => total + rule.points, 0);

  return {
    riskScore: score,
    riskLevel: levelFor(score),
    riskReasons: hits.map((rule) => rule.reason),
  };
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/derive/riskScore.test.js`
Expected: PASS — 15 tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/derive/
git commit -m "Score device risk additively with a reason per charge"
```

---

## Task 10: `deriveDevice` — the orchestrator

Implements spec §5, and assembles §8.1–§8.7 into one record.

**Files:**
- Create: `src/features/devices/derive/deriveDevice.js`
- Test: `src/features/devices/derive/deriveDevice.test.js`

**Interfaces:**
- Consumes: `parseReport` (Task 2), every parser from Task 3, `deriveRam`, `deriveStorage`, `deriveCpu`, `deriveIdentity`, `deriveHealth`, `riskScore` (Tasks 4–9).
- Produces: `deriveDevice({ text, fileName, lastModified }): DeviceRecord`

`DeviceRecord` is the shape every later task consumes. Its keys match the
SharePoint `StaticName`s in spec §9.2 exactly, in `camelCase`:

```
computerName owner ownerSource department deviceType deviceTypeConfident
computerModel motherboardVendor motherboardModel anydeskId
scannedOn importedOn sourceFileName
windowsVersion windowsMajor windowsEdition osSupported
cpuModel cpuVendor cpuGeneration cpuAgeBand
installedRamGB reportedRamGB ramDiscrepancy ramType ramSpeedMhz
ramSlotsUsed ramSlotsTotal ramUpgradable ramSlotInfoRaw
storageTotalGB driveCount storageType hasHdd storageDrivesRaw
antivirusStatus antivirusStatusRaw antivirusProducts avProtected
networkType ssid ipAddress ipAssignment
gpuList monitorCount monitorsRaw
microsoftOffice adobeProducts mappedDrives serverFolders serverCredentials
mailboxCount archiveCount emailDataFiles
riskScore riskLevel riskReasons scanComplete remarks
unknownLabels rawReport
```

`scannedOn` and `importedOn` are epoch milliseconds; conversion to ISO happens
at the SharePoint boundary, not here.

- [ ] **Step 1: Write the failing test**

`src/features/devices/derive/deriveDevice.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { deriveDevice } from './deriveDevice.js';

const load = (name) => ({
  text: readFileSync(fileURLToPath(new URL(`../__fixtures__/${name}`, import.meta.url)), 'utf8'),
  fileName: name,
  lastModified: Date.UTC(2026, 7, 19, 1, 18),
});

describe('deriveDevice — ASHRAF-PC', () => {
  const device = deriveDevice(load('ASHRAF-PC_.txt'));

  it('reads identity', () => {
    expect(device.computerName).toBe('ASHRAF-PC');
    expect(device.owner).toBe('Ashraf');
    expect(device.ownerSource).toBe('Server credential');
    expect(device.deviceType).toBe('Laptop');
    expect(device.department).toBe(null);
  });

  it('reads specs', () => {
    expect(device.installedRamGB).toBe(8);
    expect(device.storageTotalGB).toBe(477);
    expect(device.storageType).toBe('SSD only');
    expect(device.cpuGeneration).toBe('13');
    expect(device.windowsMajor).toBe(11);
  });

  it('counts real GPUs and monitors, not the virtual ones', () => {
    expect(device.gpuList).toEqual(['Intel(R) Iris(R) Xe Graphics']);
    expect(device.monitorCount).toBe(1);
  });

  it('counts mailboxes and archives separately', () => {
    expect(device.mailboxCount).toBe(1);
    expect(device.archiveCount).toBe(2);
  });

  it('scores it Watch on its 8 GB alone', () => {
    expect(device.riskScore).toBe(15);
    expect(device.riskLevel).toBe('Watch');
  });

  it('keeps the raw report for later re-derivation', () => {
    expect(device.rawReport).toContain('KBG50ZNV512G KIOXIA');
  });

  it('carries both timestamps through', () => {
    expect(device.scannedOn).toBe(Date.UTC(2026, 7, 19, 1, 18));
    expect(typeof device.importedOn).toBe('number');
    expect(device.sourceFileName).toBe('ASHRAF-PC_.txt');
  });
});

describe('deriveDevice — the awkward machines', () => {
  it('marks the failed CARMEN-HP scan incomplete with no risk score', () => {
    const device = deriveDevice(load('CARMEN-HP_.txt'));
    expect(device.scanComplete).toBe(false);
    expect(device.riskScore).toBe(null);
    expect(device.riskLevel).toBe('Unknown');
    expect(device.computerName).toBe('CARMEN-HP');
  });

  it('scores DESKTOP-8SBR420 Critical', () => {
    const device = deriveDevice(load('[STOCKYARDF1] DESKTOP-8SBR420_.txt'));
    expect(device.department).toBe('STOCKYARDF1');
    expect(device.deviceType).toBe('Desktop');
    expect(device.installedRamGB).toBe(2);
    expect(device.storageType).toBe('Mixed');
    expect(device.riskLevel).toBe('Critical');
  });

  it('reports the RAM discrepancy on EVONNE-HP and flags the free slot', () => {
    const device = deriveDevice(load('[FINANCE] EVONNE-HP_.txt'));
    expect(device.installedRamGB).toBe(8);
    expect(device.ramSlotsUsed).toBe(1);
    expect(device.ramSlotsTotal).toBe(2);
    expect(device.ramUpgradable).toBe(true);
  });

  it('collapses the 22 duplicated antivirus entries on AMIR-HP', () => {
    const device = deriveDevice(load('[ENGINEERING] AMIR-HP_.txt'));
    expect(device.antivirusProducts).toHaveLength(3);
    expect(device.avProtected).toBe(true);
    expect(device.hasHdd).toBe(true);
    expect(device.riskLevel).toBe('High');
  });

  it('reads Core Ultra on PMWL034', () => {
    const device = deriveDevice(load('PMWL034_.txt'));
    expect(device.cpuGeneration).toBe('Ultra 1');
    expect(device.cpuAgeBand).toBe('Current');
    expect(device.riskLevel).toBe('OK');
  });

  it('reads the person out of a combined department bracket', () => {
    const device = deriveDevice(load('[QAQC FAIRUS]HPFL05_.txt'));
    expect(device.department).toBe('QAQC');
    expect(device.owner).toBe('Fairus');
    expect(device.riskLevel).toBe('Critical');
  });

  it('counts the mapped drives on PGCHAN-HP', () => {
    const device = deriveDevice(load('[SALES] PGCHAN-HP_.txt'));
    expect(device.mappedDrives).toBe(12);
    expect(device.archiveCount).toBe(6);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/derive/deriveDevice.test.js`
Expected: FAIL — `Failed to resolve import "./deriveDevice.js"`

- [ ] **Step 3: Write `deriveDevice.js`**

```js
import { parseReport } from '../parse/parseReport.js';
import { cleanValue } from '../parse/placeholders.js';
import {
  parsePairs, parseNetwork, parseOffice, parseGpus, parseMonitors, parseMailFiles,
} from '../parse/parseValues.js';
import { deriveRam } from './deriveRam.js';
import { deriveStorage } from './deriveStorage.js';
import { deriveCpu } from './deriveCpu.js';
import { deriveIdentity } from './deriveIdentity.js';
import { deriveHealth } from './deriveHealth.js';
import { riskScore } from './riskScore.js';

const firstOrNull = (lines) => (lines?.length ? cleanValue(lines[0]) : null);
const joinLines = (lines) => (lines?.length ? lines.join('\n') : null);

export function deriveDevice({ text, fileName, lastModified }) {
  const { fields, unknownLabels } = parseReport(text);

  const identity = deriveIdentity(fields, fileName);
  const ram = deriveRam(fields['RAM Slot Info'] ?? [], fields['Total RAM'] ?? []);
  const storage = deriveStorage(fields['Storage Drives'] ?? []);
  const cpu = deriveCpu(fields['Processor'] ?? [], ram.ramType);
  const health = deriveHealth(fields);

  const [motherboard] = parsePairs(fields['Motherboard'] ?? []);
  const network = parseNetwork(fields['Network Information'] ?? []);
  const mail = parseMailFiles(fields['Email data files found Active or Inactive account'] ?? []);
  const serverFolders = parsePairs(fields['Server folder'] ?? []);

  const base = {
    ...identity,
    ...ram,
    ...storage,
    ...cpu,
    ...health,

    computerModel: firstOrNull(fields['Computer Model']),
    motherboardVendor: motherboard?.left ?? null,
    motherboardModel: motherboard?.right ?? null,
    anydeskId: firstOrNull(fields['Anydesk']),
    remarks: joinLines(fields['Remarks']),

    scannedOn: lastModified,
    importedOn: Date.now(),
    sourceFileName: fileName,

    networkType: network?.connection ?? null,
    ssid: network?.ssid ?? null,
    ipAddress: network?.ip ?? null,
    ipAssignment: network?.assignment ?? 'Unknown',

    gpuList: parseGpus(fields['GPU'] ?? []),
    monitorCount: parseMonitors(fields['Monitor'] ?? []).length,
    monitorsRaw: joinLines(fields['Monitor']),

    microsoftOffice: parseOffice(fields['Microsoft Office'] ?? []),
    adobeProducts: parsePairs(fields['Adobe'] ?? [])
      .filter((entry) => entry.left)
      .map((entry) => (entry.right ? `${entry.left} ${entry.right}` : entry.left)),
    mappedDrives: serverFolders.filter((entry) => entry.left).length,
    serverFolders: joinLines(fields['Server folder']),
    serverCredentials: joinLines(fields['PMW Server and credentials']),

    mailboxCount: mail.filter((entry) => entry.kind === 'mailbox').length,
    archiveCount: mail.filter((entry) => entry.kind === 'archive').length,
    emailDataFiles: joinLines(fields['Email data files found Active or Inactive account']),

    ramSlotInfoRaw: joinLines(fields['RAM Slot Info']),
    storageDrivesRaw: joinLines(fields['Storage Drives']),

    unknownLabels,
    rawReport: text,
  };

  return { ...base, ...riskScore(base) };
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/derive/deriveDevice.test.js`
Expected: PASS — 13 tests.

If a fixture assertion fails, **check the fixture before changing the code** —
these numbers were read off the real files, so a mismatch means a derivation
bug, not a bad expectation.

- [ ] **Step 5: Run the whole suite**

Run: `npm test`
Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add src/features/devices/
git commit -m "Assemble parsed device fields into one typed record"
```

---

## Task 11: Twelve-hour Malaysia time

Implements spec §9.4.

**Files:**
- Modify: `src/features/datastudio/time/malaysiaTime.js` (the `formatMYT` function)
- Modify: `src/features/datastudio/time/malaysiaTime.test.js` (append a describe block)

**Interfaces:**
- Produces: `formatMYT(epochMs, style)` gains two styles — `'datetime12'` → `19/08/2026 09:18 AM`, `'time12'` → `09:18 AM`. Existing styles `'date'`, `'time'`, `'datetime'` are unchanged.

**Do not touch `getPartsMYT`.** It pins `hourCycle: 'h23'` behind a comment
explaining that an explicit `hour12` nullifies `hourCycle` entirely. The
twelve-hour path follows the same discipline in the opposite direction: it pins
`hourCycle: 'h12'` and never passes `hour12`. `h12` renders midnight as `12 AM`;
`h11` would render it as `0 AM`, which is the mirror of the bug the existing
comment guards against.

- [ ] **Step 1: Write the failing test**

Append to `src/features/datastudio/time/malaysiaTime.test.js`:

```js
describe('formatMYT — twelve-hour styles', () => {
  // 19 Aug 2026, 01:18 UTC = 09:18 MYT
  const morning = Date.UTC(2026, 7, 19, 1, 18);
  // 19 Aug 2026, 09:04 UTC = 17:04 MYT
  const evening = Date.UTC(2026, 7, 19, 9, 4);
  // 18 Aug 2026, 16:00 UTC = 19 Aug 2026, 00:00 MYT
  const midnight = Date.UTC(2026, 7, 18, 16, 0);
  // 19 Aug 2026, 04:00 UTC = 12:00 MYT
  const noon = Date.UTC(2026, 7, 19, 4, 0);

  it('renders a morning time with AM', () => {
    expect(formatMYT(morning, 'datetime12')).toBe('19/08/2026 09:18 AM');
  });

  it('renders an afternoon time with PM', () => {
    expect(formatMYT(evening, 'datetime12')).toBe('19/08/2026 05:04 PM');
  });

  it('renders midnight as 12 AM, not 0 AM', () => {
    expect(formatMYT(midnight, 'datetime12')).toBe('19/08/2026 12:00 AM');
  });

  it('renders noon as 12 PM', () => {
    expect(formatMYT(noon, 'time12')).toBe('12:00 PM');
  });

  it('leaves the twenty-four hour styles untouched', () => {
    expect(formatMYT(evening, 'datetime')).toBe('19/08/2026 17:04');
    expect(formatMYT(midnight, 'time')).toBe('00:00');
  });

  it('returns the em dash for an unparseable value', () => {
    expect(formatMYT(NaN, 'datetime12')).toBe('—');
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/datastudio/time/malaysiaTime.test.js`
Expected: FAIL — `datetime12` falls through to the default branch and returns
the 24-hour string `19/08/2026 09:18`.

- [ ] **Step 3: Add the twelve-hour parts helper and the two styles**

In `src/features/datastudio/time/malaysiaTime.js`, add below `getPartsMYT`:

```js
/**
 * The twelve-hour twin of getPartsMYT. Same discipline as its sibling: pin the
 * hour cycle and never pass `hour12`, because an explicit hour12 nullifies
 * hourCycle. 'h12' is the 1-12 cycle that renders midnight as 12 AM; 'h11'
 * would render it as 0 AM.
 */
function getParts12MYT(epochMs, options) {
  const formatter = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kuala_Lumpur',
    hourCycle: 'h12',
    ...options,
  });
  const byType = {};
  for (const part of formatter.formatToParts(new Date(epochMs))) {
    byType[part.type] = part.value;
  }
  return byType;
}
```

Then, inside `formatMYT`, add the twelve-hour time builder alongside `timePart`
and route the two new styles before the existing returns:

```js
  const timePart12 = () => {
    const { hour, minute, dayPeriod } = getParts12MYT(epochMs, {
      hour: '2-digit',
      minute: '2-digit',
    });
    return `${hour}:${minute} ${dayPeriod.toUpperCase()}`;
  };

  if (style === 'date') return datePart();
  if (style === 'time') return timePart();
  if (style === 'time12') return timePart12();
  if (style === 'datetime12') return `${datePart()} ${timePart12()}`;
  return `${datePart()} ${timePart()}`;
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/datastudio/time/malaysiaTime.test.js`
Expected: PASS — the 33 existing tests plus 6 new ones.

If `dayPeriod` comes back as `am`/`pm` on this engine, the `.toUpperCase()`
already handles it. If it comes back with a narrow no-break space before it in
some ICU build, the assertion will show it — normalise with
`.replace(/ /g, ' ')` rather than loosening the test.

- [ ] **Step 5: Commit**

```bash
git add src/features/datastudio/time/
git commit -m "Add twelve-hour Malaysia time styles for AM/PM display"
```

---

## Task 12: `importFiles` — File[] to records

Implements spec §10.1 (drop stage) and §11 (batch duplicates).

**Files:**
- Create: `src/features/devices/importFiles.js`
- Test: `src/features/devices/importFiles.test.js`

**Interfaces:**
- Consumes: `deriveDevice` (Task 10).
- Produces:
  - `readTextFile(file): Promise<string>` — thin `file.text()` wrapper so tests can pass a plain object
  - `importFiles(files): Promise<{ devices: DeviceRecord[], rejected: {fileName, reason}[] }>`

**Duplicate rule:** within one batch, two files naming the same computer keep
the one with the newer `scannedOn`; the loser is rejected with a reason naming
the winner.

- [ ] **Step 1: Write the failing test**

`src/features/devices/importFiles.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { importFiles } from './importFiles.js';

const fixture = (name) =>
  readFileSync(fileURLToPath(new URL(`./__fixtures__/${name}`, import.meta.url)), 'utf8');

/** Minimal stand-in for a browser File: name, lastModified, text(). */
const fakeFile = (name, lastModified, text = fixture(name)) => ({
  name, lastModified, text: async () => text,
});

describe('importFiles', () => {
  it('parses several files into records', async () => {
    const result = await importFiles([
      fakeFile('ASHRAF-PC_.txt', 1_760_000_000_000),
      fakeFile('PMWL034_.txt', 1_760_000_000_000),
    ]);
    expect(result.devices.map((d) => d.computerName)).toEqual(['ASHRAF-PC', 'PMWL034']);
    expect(result.rejected).toEqual([]);
  });

  it('rejects a file that is not a .txt', async () => {
    const result = await importFiles([fakeFile('report.pdf', 0, 'nonsense')]);
    expect(result.devices).toEqual([]);
    expect(result.rejected).toEqual([
      { fileName: 'report.pdf', reason: 'Not a .txt file' },
    ]);
  });

  it('rejects a .txt that is not a device report', async () => {
    const result = await importFiles([
      fakeFile('invoice.txt', 0, 'Dear team,\n\nPlease find the invoice attached.\n'),
    ]);
    expect(result.rejected).toEqual([
      { fileName: 'invoice.txt', reason: 'Not a device report — no known fields found' },
    ]);
  });

  it('keeps the newer of two files naming the same computer', async () => {
    const older = fakeFile('ASHRAF-PC_.txt', 1_700_000_000_000);
    const newer = fakeFile('[IT] ASHRAF-PC_.txt', 1_760_000_000_000, fixture('ASHRAF-PC_.txt'));
    const result = await importFiles([older, newer]);

    expect(result.devices).toHaveLength(1);
    expect(result.devices[0].sourceFileName).toBe('[IT] ASHRAF-PC_.txt');
    expect(result.rejected).toEqual([
      {
        fileName: 'ASHRAF-PC_.txt',
        reason: 'Duplicate of ASHRAF-PC — kept the newer scan from [IT] ASHRAF-PC_.txt',
      },
    ]);
  });

  it('reports a file it could not read without losing the rest of the batch', async () => {
    const broken = {
      name: 'broken.txt', lastModified: 0,
      text: async () => { throw new Error('disk gone'); },
    };
    const result = await importFiles([broken, fakeFile('PMWL034_.txt', 0)]);

    expect(result.devices).toHaveLength(1);
    expect(result.rejected).toEqual([
      { fileName: 'broken.txt', reason: 'Could not read the file: disk gone' },
    ]);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/importFiles.test.js`
Expected: FAIL — `Failed to resolve import "./importFiles.js"`

- [ ] **Step 3: Write `importFiles.js`**

```js
import { deriveDevice } from './derive/deriveDevice.js';
import { parseReport } from './parse/parseReport.js';

export function readTextFile(file) {
  return file.text();
}

export async function importFiles(files) {
  const devices = [];
  const rejected = [];

  for (const file of files) {
    if (!/\.txt$/i.test(file.name)) {
      rejected.push({ fileName: file.name, reason: 'Not a .txt file' });
      continue;
    }

    let text;
    try {
      text = await readTextFile(file);
    } catch (error) {
      rejected.push({ fileName: file.name, reason: `Could not read the file: ${error.message}` });
      continue;
    }

    // Checked before deriving so an unrelated .txt is named as such rather
    // than imported as a machine with every field empty.
    if (!parseReport(text).isReport) {
      rejected.push({
        fileName: file.name,
        reason: 'Not a device report — no known fields found',
      });
      continue;
    }

    devices.push(deriveDevice({ text, fileName: file.name, lastModified: file.lastModified }));
  }

  return dedupe(devices, rejected);
}

function dedupe(devices, rejected) {
  const byName = new Map();

  for (const device of devices) {
    const key = (device.computerName ?? device.sourceFileName).toLowerCase();
    const existing = byName.get(key);

    if (!existing) {
      byName.set(key, device);
      continue;
    }

    const [keep, drop] = device.scannedOn > existing.scannedOn
      ? [device, existing]
      : [existing, device];

    byName.set(key, keep);
    rejected.push({
      fileName: drop.sourceFileName,
      reason: `Duplicate of ${keep.computerName} — kept the newer scan from ${keep.sourceFileName}`,
    });
  }

  return { devices: [...byName.values()], rejected };
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/importFiles.test.js`
Expected: PASS — 5 tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/
git commit -m "Turn dropped files into device records, rejecting what is not one"
```

---

# PHASE 3 — Drop and review, with no SharePoint

At the end of this phase the section is fully usable offline: drop the 17 real
files, see them parsed, correct what was guessed. Nothing writes anywhere.

---

## Task 13: The route, the dropzone, and a parsed summary

Implements spec §10.1 (drop stage), §6 (file layout), §4.6.

**Files:**
- Create: `src/pages/DevicesPage.jsx`
- Create: `src/features/devices/ui/DropZone.jsx`
- Create: `src/styles/devices.css`
- Modify: `src/App.jsx` (add the route)
- Modify: `src/components/AppShell.jsx` (add the nav item)
- Modify: `src/main.jsx` (import `devices.css` after `shell.css`)
- Modify: `src/components/ui/Icons.jsx` (add `HardDrive`, `Cpu`, `MemoryStick`)

**Interfaces:**
- Consumes: `importFiles` (Task 12), `Card`/`EmptyState` from `src/components/ui/Surfaces.jsx`, `Button` from `src/components/ui/Button.jsx`.
- Produces: `<DropZone onFiles={(FileList|File[]) => void} busy={boolean} />`

**Existing APIs you will need, exactly as they are:**
- `AppShell` takes `{ title, subtitle, actions, search, children }`.
- `NAV_ITEMS` in `AppShell.jsx` is `{ to, label, icon }`; icons come from `src/components/ui/Icons.jsx`, which already exports `Laptop`, `AlertTriangle`, `ShieldCheck`, `Download`, `Inbox`, `Building`, `Clock`, `Filter`, `Check`.
- `Button` takes `{ variant, size, icon, ...props }`.

- [ ] **Step 1: Add the three missing icons**

In `src/components/ui/Icons.jsx`, following the file's existing 24px stroke-grid
style, add:

```jsx
export function HardDrive({ size = 24, ...props }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor"
      strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...props}>
      <line x1="22" y1="12" x2="2" y2="12" />
      <path d="M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z" />
      <line x1="6" y1="16" x2="6.01" y2="16" />
      <line x1="10" y1="16" x2="10.01" y2="16" />
    </svg>
  );
}

export function Cpu({ size = 24, ...props }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor"
      strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...props}>
      <rect x="4" y="4" width="16" height="16" rx="2" />
      <rect x="9" y="9" width="6" height="6" />
      <path d="M9 2v2M15 2v2M9 20v2M15 20v2M2 9h2M2 15h2M20 9h2M20 15h2" />
    </svg>
  );
}

export function MemoryStick({ size = 24, ...props }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor"
      strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...props}>
      <path d="M6 19v-3M10 19v-3M14 19v-3M18 19v-3" />
      <path d="M8 11V9M16 11V9M12 11V9" />
      <path d="M2 15h20" />
      <path d="M4 15V7a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v8" />
    </svg>
  );
}
```

- [ ] **Step 2: Write `DropZone.jsx`**

```jsx
import { useCallback, useRef, useState } from 'react';
import { Inbox } from '../../../components/ui/Icons';
import Button from '../../../components/ui/Button';

/**
 * Drag and drop OR a file picker. The picker is not optional: dragging an
 * attachment straight out of Outlook does not reliably produce a File, and
 * that is exactly where these reports arrive from.
 */
export default function DropZone({ onFiles, busy }) {
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef(null);

  const stop = (event) => {
    event.preventDefault();
    event.stopPropagation();
  };

  const handleDrop = useCallback(
    (event) => {
      stop(event);
      setDragging(false);
      const files = [...(event.dataTransfer?.files ?? [])];
      if (files.length) onFiles(files);
    },
    [onFiles],
  );

  return (
    <div
      className={`dz${dragging ? ' dz-active' : ''}${busy ? ' dz-busy' : ''}`}
      onDragEnter={(e) => { stop(e); setDragging(true); }}
      onDragOver={stop}
      onDragLeave={(e) => { stop(e); setDragging(false); }}
      onDrop={handleDrop}
    >
      <Inbox size={28} className="dz-icon" />
      <p className="dz-title">Drop device report files here</p>
      <p className="dz-hint">
        The <code>.txt</code> reports the scan script writes. Drop as many as you like —
        one row per file. Nothing is saved until you review it.
      </p>
      <Button
        variant="secondary"
        onClick={() => inputRef.current?.click()}
        disabled={busy}
      >
        Choose files
      </Button>
      <input
        ref={inputRef}
        type="file"
        multiple
        accept=".txt,text/plain"
        className="dz-input"
        onChange={(event) => {
          const files = [...(event.target.files ?? [])];
          if (files.length) onFiles(files);
          event.target.value = '';
        }}
      />
    </div>
  );
}
```

- [ ] **Step 3: Write `DevicesPage.jsx`**

The page owns the stage. Stage is component state, not a route, so a half-done
review is not lost to a stray back button.

```jsx
import { useCallback, useState } from 'react';
import AppShell from '../components/AppShell';
import { Card, EmptyState } from '../components/ui/Surfaces';
import DropZone from '../features/devices/ui/DropZone';
import { importFiles } from '../features/devices/importFiles';

export default function DevicesPage() {
  const [stage, setStage] = useState('drop');
  const [devices, setDevices] = useState([]);
  const [rejected, setRejected] = useState([]);
  const [busy, setBusy] = useState(false);

  const handleFiles = useCallback(async (files) => {
    setBusy(true);
    try {
      const result = await importFiles(files);
      setDevices(result.devices);
      setRejected(result.rejected);
      setStage(result.devices.length ? 'review' : 'drop');
    } finally {
      setBusy(false);
    }
  }, []);

  const reset = () => {
    setDevices([]);
    setRejected([]);
    setStage('drop');
  };

  return (
    <AppShell
      title="Device list"
      subtitle="Import machine scan reports and keep the fleet register current"
    >
      {stage === 'drop' && (
        <Card className="dz-card">
          <DropZone onFiles={handleFiles} busy={busy} />
          {rejected.length > 0 && (
            <ul className="dz-rejected">
              {rejected.map((item) => (
                <li key={item.fileName}>
                  <strong>{item.fileName}</strong> — {item.reason}
                </li>
              ))}
            </ul>
          )}
        </Card>
      )}

      {stage === 'review' && (
        <Card>
          <div className="review-head">
            <p className="review-summary">{devices.length} report(s) parsed</p>
            <button type="button" className="ui-btn ui-btn-sm ui-btn-ghost" onClick={reset}>
              Start over
            </button>
          </div>
          {devices.length === 0 && <EmptyState>Nothing to review.</EmptyState>}
          <ul className="review-preview">
            {devices.map((device) => (
              <li key={device.sourceFileName}>
                {device.computerName} — {device.owner ?? 'no owner'} —{' '}
                {device.installedRamGB ?? '?'} GB — {device.riskLevel}
              </li>
            ))}
          </ul>
        </Card>
      )}
    </AppShell>
  );
}
```

The `review` stage is replaced wholesale by the real grid in Task 14. It exists
here so this task has something to verify.

- [ ] **Step 4: Wire the route, the nav item and the stylesheet**

In `src/App.jsx`, add alongside the existing routes:

```jsx
<Route path="/devices" element={<DevicesPage />} />
```

In `src/components/AppShell.jsx`, add to `NAV_ITEMS` after the asset checklist
entry, importing `Laptop` from `./ui/Icons` if it is not already imported:

```jsx
  { to: '/devices', label: 'Device list', icon: Laptop },
```

In `src/main.jsx`, add the import **after** `styles/shell.css`:

```js
import './styles/devices.css';
```

- [ ] **Step 5: Write `devices.css`**

Tokens only — no hex values. Mobile-first, `min-width` breakpoints.

```css
/* Device list — drop, review, register and dashboard. Loaded after shell.css. */

.dz-card { padding: 0; }

.dz {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 10px;
  padding: 40px 20px;
  border: 1px dashed var(--it-line);
  border-radius: var(--it-radius);
  background: var(--it-panel);
  text-align: center;
  transition: border-color 120ms ease, background 120ms ease;
}

.dz-active {
  border-color: var(--it-brand);
  background: var(--it-brand-wash);
}

.dz-busy { opacity: 0.6; pointer-events: none; }

.dz-icon { color: var(--it-ink-soft); }

.dz-title {
  margin: 0;
  font-weight: 600;
  color: var(--it-ink);
}

.dz-hint {
  margin: 0;
  max-width: 46ch;
  font-size: 0.85rem;
  line-height: 1.5;
  color: var(--it-ink-soft);
}

.dz-input { display: none; }

.dz-rejected {
  margin: 0;
  padding: 12px 20px 20px;
  list-style: none;
  font-size: 0.82rem;
  color: var(--it-danger);
}

.dz-rejected li { padding: 2px 0; }

.review-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  margin-bottom: 12px;
}

.review-summary { margin: 0; font-weight: 600; color: var(--it-ink); }

.review-preview {
  margin: 0;
  padding-left: 18px;
  font-size: 0.85rem;
  color: var(--it-ink-soft);
}

@media (prefers-reduced-motion: reduce) {
  .dz { transition: none; }
}
```

- [ ] **Step 6: Verify in the browser**

Start the preview and drop the real files:

1. `preview_start` with the dev server.
2. Navigate to `/devices`.
3. Drop all 17 files from `~/Downloads`.
4. Confirm: 17 rows listed, `CARMEN-HP` shows `Unknown` risk, `DESKTOP-8SBR420`
   shows `Critical`, and dropping a non-`.txt` file names it in the rejected list.
5. `read_console_messages` — expect no errors.

- [ ] **Step 7: Lint**

Run: `npm run lint`
Expected: no new errors in `DevicesPage.jsx`, `DropZone.jsx`, `Icons.jsx`,
`App.jsx`, `AppShell.jsx`. The four pre-existing failures stay.

- [ ] **Step 8: Commit**

```bash
git add src/ && git commit -m "Add the /devices route with a working report dropzone"
```

---

## Task 14: `ReviewGrid` — the editable review table

Implements spec §10.1 (review stage).

**Files:**
- Create: `src/features/devices/ui/ReviewGrid.jsx`
- Create: `src/features/devices/reviewIssues.js`
- Test: `src/features/devices/reviewIssues.test.js`
- Modify: `src/pages/DevicesPage.jsx` (replace the placeholder review stage)
- Modify: `src/styles/devices.css` (append the grid styles)

**Interfaces:**
- Consumes: `formatMYT` from `src/features/datastudio/time/malaysiaTime.js` (Task 11), the `DeviceRecord` shape (Task 10).
- Produces:
  - `issuesFor(device): string[]` — the problems worth sorting to the top
  - `sortForReview(devices): DeviceRecord[]` — problems first, then by computer name
  - `<ReviewGrid devices editable onChange onToggleRow excluded />`

The derived fields — `owner`, `department`, `deviceType`, `scannedOn` — are the
only editable ones. Everything else came verbatim out of the file and editing it
would put the register out of step with the report that produced it.

- [ ] **Step 1: Write the failing test**

`src/features/devices/reviewIssues.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { issuesFor, sortForReview } from './reviewIssues.js';

const clean = {
  computerName: 'PC1', scanComplete: true, deviceType: 'Laptop',
  ramDiscrepancy: false, unknownLabels: [], owner: 'Ali',
};

describe('issuesFor', () => {
  it('finds nothing wrong with a clean record', () => {
    expect(issuesFor(clean)).toEqual([]);
  });

  it('reports an incomplete scan', () => {
    expect(issuesFor({ ...clean, scanComplete: false }))
      .toEqual(['Scan incomplete — most fields are empty']);
  });

  it('reports an unresolved device type', () => {
    expect(issuesFor({ ...clean, deviceType: 'Unknown' }))
      .toEqual(['Device type could not be determined']);
  });

  it('explains the RAM discrepancy rather than just flagging it', () => {
    expect(issuesFor({ ...clean, ramDiscrepancy: true, installedRamGB: 16, reportedRamGB: 15 }))
      .toEqual(['Reports 15 GB usable of 16 GB installed — the GPU reserves the rest']);
  });

  it('reports an unknown label by name', () => {
    expect(issuesFor({ ...clean, unknownLabels: [{ label: 'BitLocker Status', value: 'On' }] }))
      .toEqual(['New field found in the report: BitLocker Status']);
  });

  it('reports a missing owner', () => {
    expect(issuesFor({ ...clean, owner: null })).toEqual(['No owner could be resolved']);
  });
});

describe('sortForReview', () => {
  it('puts rows with problems first, then sorts by name', () => {
    const rows = [
      { ...clean, computerName: 'BBB' },
      { ...clean, computerName: 'AAA' },
      { ...clean, computerName: 'ZZZ', scanComplete: false },
    ];
    expect(sortForReview(rows).map((r) => r.computerName)).toEqual(['ZZZ', 'AAA', 'BBB']);
  });

  it('does not mutate the input', () => {
    const rows = [{ ...clean, computerName: 'B' }, { ...clean, computerName: 'A' }];
    sortForReview(rows);
    expect(rows.map((r) => r.computerName)).toEqual(['B', 'A']);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/reviewIssues.test.js`
Expected: FAIL — `Failed to resolve import "./reviewIssues.js"`

- [ ] **Step 3: Write `reviewIssues.js`**

```js
export function issuesFor(device) {
  const issues = [];

  if (device.scanComplete === false) issues.push('Scan incomplete — most fields are empty');
  if (device.deviceType === 'Unknown') issues.push('Device type could not be determined');

  if (device.ramDiscrepancy) {
    issues.push(
      `Reports ${device.reportedRamGB} GB usable of ${device.installedRamGB} GB installed ` +
        '— the GPU reserves the rest',
    );
  }

  for (const unknown of device.unknownLabels ?? []) {
    issues.push(`New field found in the report: ${unknown.label}`);
  }

  if (!device.owner) issues.push('No owner could be resolved');

  return issues;
}

export function sortForReview(devices) {
  return [...devices].sort((a, b) => {
    const problems = issuesFor(b).length > 0 ? 1 : 0;
    const mine = issuesFor(a).length > 0 ? 1 : 0;
    if (problems !== mine) return problems - mine;
    return (a.computerName ?? '').localeCompare(b.computerName ?? '');
  });
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/reviewIssues.test.js`
Expected: PASS — 8 tests.

- [ ] **Step 5: Write `ReviewGrid.jsx`**

```jsx
import { formatMYT } from '../../datastudio/time/malaysiaTime';
import { issuesFor, sortForReview } from '../reviewIssues';

const COLUMNS = [
  { key: 'computerName', label: 'Computer' },
  { key: 'owner', label: 'Owner', editable: true },
  { key: 'department', label: 'Department', editable: true },
  { key: 'deviceType', label: 'Type', editable: true, options: ['Laptop', 'Desktop', 'Unknown'] },
  { key: 'computerModel', label: 'Model' },
  { key: 'cpuModel', label: 'CPU' },
  { key: 'installedRamGB', label: 'RAM (GB)' },
  { key: 'storageTotalGB', label: 'Storage (GB)' },
  { key: 'storageType', label: 'Disks' },
  { key: 'windowsVersion', label: 'Windows' },
  { key: 'antivirusStatus', label: 'Antivirus' },
  { key: 'riskLevel', label: 'Risk' },
];

export default function ReviewGrid({ devices, excluded, onChange, onToggleRow }) {
  const rows = sortForReview(devices);

  return (
    <div className="rg-scroll">
      <table className="rg">
        <thead>
          <tr>
            <th className="rg-check"><span className="sr-only">Include</span></th>
            {COLUMNS.map((column) => (
              <th key={column.key}>{column.label}</th>
            ))}
            <th>Scanned</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((device) => {
            const issues = issuesFor(device);
            const id = device.sourceFileName;
            const isExcluded = excluded.has(id);

            return (
              <tr key={id} className={issues.length ? 'rg-flagged' : undefined}>
                <td className="rg-check">
                  <input
                    type="checkbox"
                    checked={!isExcluded}
                    onChange={() => onToggleRow(id)}
                    aria-label={`Include ${device.computerName}`}
                  />
                </td>

                {COLUMNS.map((column) => {
                  const value = device[column.key];

                  if (!column.editable) {
                    return (
                      <td key={column.key} className={column.key === 'riskLevel'
                        ? `rg-risk rg-risk-${String(value).toLowerCase()}` : undefined}>
                        {value ?? '—'}
                      </td>
                    );
                  }

                  return (
                    <td key={column.key} className="rg-editable">
                      {column.options ? (
                        <select
                          value={value ?? 'Unknown'}
                          onChange={(event) => onChange(id, column.key, event.target.value)}
                        >
                          {column.options.map((option) => (
                            <option key={option} value={option}>{option}</option>
                          ))}
                        </select>
                      ) : (
                        <input
                          type="text"
                          value={value ?? ''}
                          placeholder="—"
                          onChange={(event) =>
                            onChange(id, column.key, event.target.value || null)}
                        />
                      )}
                    </td>
                  );
                })}

                <td title="Malaysia time">{formatMYT(device.scannedOn, 'datetime12')}</td>
              </tr>
            );
          })}
        </tbody>
      </table>

      <ul className="rg-issues">
        {rows.flatMap((device) =>
          issuesFor(device).map((issue) => (
            <li key={`${device.sourceFileName}-${issue}`}>
              <strong>{device.computerName}</strong> — {issue}
            </li>
          )),
        )}
      </ul>
    </div>
  );
}
```

The `editable: true` flag on a column is what carries the "this was guessed"
marker, styled through `.rg-editable`. Keep `COLUMNS` module-private — exporting
a non-component from a component module drops the file out of Fast Refresh and
fails `npm run lint` (see Global Constraints).

- [ ] **Step 6: Wire it into `DevicesPage.jsx`**

Replace the placeholder review stage. Add the two pieces of state and the edit
handler:

```jsx
  const [excluded, setExcluded] = useState(new Set());
  const [edits, setEdits] = useState({});

  const merged = devices.map((device) => ({ ...device, ...(edits[device.sourceFileName] ?? {}) }));

  const handleChange = (id, key, value) =>
    setEdits((current) => ({ ...current, [id]: { ...(current[id] ?? {}), [key]: value } }));

  const handleToggleRow = (id) =>
    setExcluded((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
```

and render:

```jsx
        <ReviewGrid
          devices={merged}
          excluded={excluded}
          onChange={handleChange}
          onToggleRow={handleToggleRow}
        />
```

Edits are held separately from the parsed records so that "start over" and a
re-parse both discard them cleanly, and so the raw record still matches the file.

- [ ] **Step 7: Append the grid styles to `devices.css`**

```css
.rg-scroll { overflow-x: auto; }

.rg {
  width: 100%;
  border-collapse: collapse;
  font-size: 0.82rem;
  white-space: nowrap;
}

.rg th,
.rg td {
  padding: 8px 10px;
  border-bottom: 1px solid var(--it-line);
  text-align: left;
}

.rg th {
  font-weight: 600;
  color: var(--it-ink-soft);
  background: var(--it-panel);
  position: sticky;
  top: 0;
}

.rg-flagged { background: var(--it-brand-wash); }

.rg-check { width: 32px; }

.rg-editable input,
.rg-editable select {
  width: 100%;
  min-width: 90px;
  padding: 3px 6px;
  border: 1px dashed var(--it-brand-line);
  border-radius: 4px;
  background: transparent;
  color: var(--it-ink);
  font: inherit;
}

.rg-risk { font-weight: 600; }
.rg-risk-critical { color: var(--it-danger); }
.rg-risk-high { color: var(--it-danger); }
.rg-risk-watch { color: var(--it-accent); }
.rg-risk-ok { color: var(--it-good); }

.rg-issues {
  margin: 12px 0 0;
  padding-left: 18px;
  font-size: 0.8rem;
  color: var(--it-ink-soft);
}

.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  overflow: hidden;
  clip: rect(0 0 0 0);
  white-space: nowrap;
}
```

- [ ] **Step 8: Verify in the browser**

Drop all 17 files again. Confirm:
- Flagged rows sort to the top, `CARMEN-HP` among them.
- The Scanned column reads as `19/08/2026 09:18 AM` — Malaysia time with AM/PM.
- Editing an owner sticks, and unchecking a row greys it out.
- `read_console_messages` shows no errors.

- [ ] **Step 9: Commit**

```bash
git add src/ && git commit -m "Add the device review grid with editable derived fields"
```

---

# PHASE 4 — SharePoint

---

## Task 15: `deviceSchema` — columns and row mapping

Implements spec §9.1, §9.2, §9.3, §9.4.

**Files:**
- Create: `src/features/devices/sharepoint/deviceSchema.js`
- Test: `src/features/devices/sharepoint/deviceSchema.test.js`

**Interfaces:**
- Consumes: `formatMYT` (Task 11).
- Produces:
  - `DEVICE_LIST_NAME = 'IT Device List'`, `CHANGE_LIST_NAME = 'IT Device Changes'`
  - `DEVICE_COLUMNS`, `CHANGE_COLUMNS` — arrays of `{ StaticName, Title, kind, choices? }`
  - `TRACKED_FIELDS: string[]` — the camelCase keys that generate change rows
  - `toListItem(device): object` — camelCase record → SharePoint item body
  - `fromListItem(row): object` — SharePoint row → camelCase record

**Serialisation rules, and why:**
- Arrays (`gpuList`, `riskReasons`, `microsoftOffice`, `antivirusProducts`) join with `\n` into Note columns. Nothing downstream needs them structured once saved, and newline-joined text is what a person opening the list wants to read.
- Epoch ms becomes `toISOString()`. SharePoint stores DateTime in UTC.
- `ScannedOnMYT` / `ChangedOnMYT` are the human-readable mirrors, because SharePoint renders DateTime in the *site's* timezone, which may not be UTC+8.
- `null` is sent as `''` for text and Note, and omitted entirely for Number, Boolean, Choice and DateTime — SharePoint rejects `null` on those.

- [ ] **Step 1: Write the failing test**

`src/features/devices/sharepoint/deviceSchema.test.js`:

```js
import { describe, it, expect } from 'vitest';
import {
  DEVICE_COLUMNS, CHANGE_COLUMNS, TRACKED_FIELDS, toListItem, fromListItem,
} from './deviceSchema.js';

const device = {
  computerName: 'ASHRAF-PC', owner: 'Ashraf', ownerSource: 'Server credential',
  department: null, deviceType: 'Laptop', computerModel: 'HP Laptop 15-fd0xxx',
  installedRamGB: 8, reportedRamGB: 8, ramDiscrepancy: false, ramUpgradable: false,
  storageTotalGB: 477, storageType: 'SSD only', hasHdd: false,
  gpuList: ['Intel(R) Iris(R) Xe Graphics'],
  riskReasons: ['8 GB of RAM or less'], riskScore: 15, riskLevel: 'Watch',
  antivirusProducts: [{ product: 'Norton 360', enabled: true }],
  scanComplete: true, scannedOn: Date.UTC(2026, 7, 19, 1, 18), importedOn: Date.UTC(2026, 7, 21, 2, 0),
  sourceFileName: 'ASHRAF-PC_.txt', rawReport: 'Name:\n', unknownLabels: [],
};

describe('DEVICE_COLUMNS', () => {
  it('has no duplicate static names', () => {
    const names = DEVICE_COLUMNS.map((c) => c.StaticName);
    expect(new Set(names).size).toBe(names.length);
  });

  it('never declares Title as a column to create — it is built in', () => {
    expect(DEVICE_COLUMNS.some((c) => c.StaticName === 'Title')).toBe(false);
  });

  it('gives every choice column its choices', () => {
    for (const column of [...DEVICE_COLUMNS, ...CHANGE_COLUMNS]) {
      if (column.kind === 'choice') expect(column.choices?.length).toBeGreaterThan(0);
    }
  });
});

describe('toListItem', () => {
  const item = toListItem(device);

  it('puts the computer name in Title', () => {
    expect(item.Title).toBe('ASHRAF-PC');
  });

  it('sends dates as ISO instants', () => {
    expect(item.ScannedOn).toBe('2026-08-19T01:18:00.000Z');
  });

  it('mirrors the scan date as Malaysia time with AM/PM', () => {
    expect(item.ScannedOnMYT).toBe('19/08/2026 09:18 AM');
  });

  it('joins arrays with newlines', () => {
    expect(item.GpuList).toBe('Intel(R) Iris(R) Xe Graphics');
    expect(item.RiskReasons).toBe('8 GB of RAM or less');
  });

  it('renders antivirus products readably', () => {
    expect(item.AntivirusProducts).toBe('Norton 360 | Enabled');
  });

  it('sends empty string for a null text column', () => {
    expect(item.Department).toBe('');
  });

  it('omits a null number rather than sending null', () => {
    const sparse = toListItem({ ...device, installedRamGB: null });
    expect('InstalledRamGB' in sparse).toBe(false);
  });

  it('omits a null choice rather than sending null', () => {
    const sparse = toListItem({ ...device, riskLevel: null });
    expect('RiskLevel' in sparse).toBe(false);
  });

  it('sends false booleans, which are not the same as absent', () => {
    expect(item.HasHdd).toBe(false);
    expect(item.ScanComplete).toBe(true);
  });
});

describe('fromListItem', () => {
  it('round-trips the fields the register needs', () => {
    const row = { ...toListItem(device), Id: 42 };
    const record = fromListItem(row);
    expect(record.id).toBe(42);
    expect(record.computerName).toBe('ASHRAF-PC');
    expect(record.installedRamGB).toBe(8);
    expect(record.hasHdd).toBe(false);
    expect(record.scannedOn).toBe(Date.UTC(2026, 7, 19, 1, 18));
    expect(record.gpuList).toEqual(['Intel(R) Iris(R) Xe Graphics']);
  });

  it('turns an absent date back into null, not NaN', () => {
    expect(fromListItem({ Title: 'X' }).scannedOn).toBe(null);
  });
});

describe('TRACKED_FIELDS', () => {
  it('tracks the hardware and health fields', () => {
    expect(TRACKED_FIELDS).toContain('installedRamGB');
    expect(TRACKED_FIELDS).toContain('riskLevel');
    expect(TRACKED_FIELDS).toContain('antivirusStatus');
  });

  it('does not track the DHCP-volatile fields', () => {
    expect(TRACKED_FIELDS).not.toContain('ipAddress');
    expect(TRACKED_FIELDS).not.toContain('ssid');
    expect(TRACKED_FIELDS).not.toContain('mappedDrives');
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/sharepoint/deviceSchema.test.js`
Expected: FAIL — `Failed to resolve import "./deviceSchema.js"`

- [ ] **Step 3: Write `deviceSchema.js`**

```js
import { formatMYT } from '../../datastudio/time/malaysiaTime.js';

export const DEVICE_LIST_NAME = 'IT Device List';
export const CHANGE_LIST_NAME = 'IT Device Changes';

const text = (StaticName, Title) => ({ StaticName, Title, kind: 'text' });
const note = (StaticName, Title) => ({ StaticName, Title, kind: 'note' });
const num = (StaticName, Title) => ({ StaticName, Title, kind: 'number' });
const bool = (StaticName, Title) => ({ StaticName, Title, kind: 'boolean' });
const date = (StaticName, Title) => ({ StaticName, Title, kind: 'datetime' });
const choice = (StaticName, Title, choices) => ({ StaticName, Title, kind: 'choice', choices });

const RISK_LEVELS = ['Critical', 'High', 'Watch', 'OK', 'Unknown'];

/** `Title` holds the computer name and is built in, so it is never created. */
export const DEVICE_COLUMNS = [
  text('Owner', 'Owner'),
  choice('OwnerSource', 'Owner Source',
    ['Name field', 'Filename', 'Server credential', 'Email', 'Manual']),
  text('Department', 'Department'),
  choice('DeviceType', 'Device Type', ['Laptop', 'Desktop', 'Unknown']),
  text('ComputerModel', 'Model'),
  text('MotherboardVendor', 'Motherboard Vendor'),
  text('MotherboardModel', 'Motherboard Model'),
  text('AnydeskId', 'AnyDesk ID'),

  date('ScannedOn', 'Scanned On'),
  date('ImportedOn', 'Imported On'),
  text('ScannedOnMYT', 'Scanned On (MYT)'),
  text('SourceFileName', 'Source File'),

  text('WindowsVersion', 'Windows Version'),
  num('WindowsMajor', 'Windows Major'),
  text('WindowsEdition', 'Windows Edition'),
  bool('OsSupported', 'OS Supported'),

  text('CpuModel', 'CPU'),
  choice('CpuVendor', 'CPU Vendor', ['Intel', 'AMD', 'Other']),
  text('CpuGeneration', 'CPU Generation'),
  choice('CpuAgeBand', 'CPU Age', ['Current', 'Aging', 'Obsolete', 'Unknown']),

  num('InstalledRamGB', 'Installed RAM (GB)'),
  num('ReportedRamGB', 'Reported RAM (GB)'),
  bool('RamDiscrepancy', 'RAM Discrepancy'),
  text('RamType', 'RAM Type'),
  num('RamSpeedMhz', 'RAM Speed (MHz)'),
  num('RamSlotsUsed', 'RAM Slots Used'),
  num('RamSlotsTotal', 'RAM Slots Total'),
  bool('RamUpgradable', 'RAM Upgradable'),
  note('RamSlotInfoRaw', 'RAM Slot Info'),

  num('StorageTotalGB', 'Storage Total (GB)'),
  num('DriveCount', 'Drive Count'),
  choice('StorageType', 'Storage Type', ['SSD only', 'Mixed', 'HDD only', 'Unknown']),
  bool('HasHdd', 'Has HDD'),
  note('StorageDrivesRaw', 'Storage Drives'),

  choice('AntivirusStatus', 'Antivirus Status',
    ['Active', 'Installed — Inactive', 'Trial', 'Not Installed', 'Unknown']),
  text('AntivirusStatusRaw', 'Antivirus Status (raw)'),
  note('AntivirusProducts', 'Antivirus Products'),
  bool('AvProtected', 'Protected'),

  text('NetworkType', 'Network'),
  text('Ssid', 'SSID'),
  text('IpAddress', 'IP Address'),
  choice('IpAssignment', 'IP Assignment', ['Dynamic', 'Static', 'Unknown']),

  note('GpuList', 'GPU'),
  num('MonitorCount', 'Monitors'),
  note('MonitorsRaw', 'Monitors (raw)'),

  note('MicrosoftOffice', 'Microsoft Office'),
  note('AdobeProducts', 'Adobe'),
  num('MappedDrives', 'Mapped Drives'),
  note('ServerFolders', 'Server Folders'),
  note('ServerCredentials', 'Server Credentials'),

  num('MailboxCount', 'Mailboxes'),
  num('ArchiveCount', 'Archives'),
  note('EmailDataFiles', 'Email Data Files'),

  num('RiskScore', 'Risk Score'),
  choice('RiskLevel', 'Risk Level', RISK_LEVELS),
  note('RiskReasons', 'Risk Reasons'),
  bool('ScanComplete', 'Scan Complete'),
  note('Remarks', 'Remarks'),
  note('ExtraFields', 'Extra Fields'),
  note('RawReport', 'Raw Report'),
];

export const CHANGE_COLUMNS = [
  text('FieldName', 'Field'),
  note('OldValue', 'Old Value'),
  note('NewValue', 'New Value'),
  date('ChangedOn', 'Changed On'),
  text('ChangedOnMYT', 'Changed On (MYT)'),
  text('ChangedBy', 'Changed By'),
  choice('ChangeType', 'Change Type', ['Added', 'Updated', 'Removed']),
];

/**
 * Only these produce change-log rows. IP address, SSID and mapped drives are
 * deliberately absent: they are DHCP-assigned or session-dependent and change
 * constantly, and logging them would bury the hardware changes that matter.
 */
export const TRACKED_FIELDS = [
  'owner', 'department', 'deviceType', 'computerModel',
  'windowsVersion', 'osSupported',
  'cpuModel', 'cpuAgeBand',
  'installedRamGB', 'ramType', 'ramSlotsUsed',
  'storageTotalGB', 'storageType',
  'antivirusStatus', 'riskLevel',
];

/** camelCase record key for a StaticName: first letter lowered. */
const keyFor = (staticName) => staticName.charAt(0).toLowerCase() + staticName.slice(1);

const serialise = (value) => {
  if (Array.isArray(value)) {
    return value
      .map((entry) =>
        entry && typeof entry === 'object' && 'product' in entry
          ? `${entry.product} | ${entry.enabled ? 'Enabled' : 'Disabled'}`
          : String(entry))
      .join('\n');
  }
  return value == null ? '' : String(value);
};

export function toListItem(device) {
  const item = { Title: device.computerName ?? '' };

  for (const column of DEVICE_COLUMNS) {
    const key = keyFor(column.StaticName);
    let value = device[key];

    if (column.StaticName === 'ScannedOnMYT') value = formatMYT(device.scannedOn, 'datetime12');
    if (column.StaticName === 'ExtraFields') {
      value = device.unknownLabels?.length ? JSON.stringify(device.unknownLabels) : null;
    }

    switch (column.kind) {
      case 'text':
      case 'note':
        // Empty string clears the column; null would be rejected.
        item[column.StaticName] = serialise(value);
        break;
      case 'number':
        if (typeof value === 'number' && Number.isFinite(value)) item[column.StaticName] = value;
        break;
      case 'boolean':
        // `false` is a real value and must survive; only null/undefined is absent.
        if (typeof value === 'boolean') item[column.StaticName] = value;
        break;
      case 'choice':
        if (value) item[column.StaticName] = String(value);
        break;
      case 'datetime':
        if (typeof value === 'number' && Number.isFinite(value)) {
          item[column.StaticName] = new Date(value).toISOString();
        }
        break;
      default:
        break;
    }
  }

  return item;
}

const ARRAY_COLUMNS = new Set(['GpuList', 'RiskReasons', 'MicrosoftOffice', 'AdobeProducts']);

export function fromListItem(row) {
  const record = { id: row.Id ?? row.ID ?? null, computerName: row.Title ?? null };

  for (const column of DEVICE_COLUMNS) {
    const key = keyFor(column.StaticName);
    const raw = row[column.StaticName];

    // An absent column reads as null for every kind — notably NOT as NaN for a
    // date, which is what `new Date(undefined).getTime()` would produce.
    if (raw === undefined || raw === null || raw === '') {
      record[key] = null;
      continue;
    }

    if (column.kind === 'datetime') record[key] = new Date(raw).getTime();
    else if (ARRAY_COLUMNS.has(column.StaticName)) record[key] = String(raw).split('\n');
    else record[key] = raw;
  }

  return record;
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/sharepoint/deviceSchema.test.js`
Expected: PASS — 15 tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/sharepoint/
git commit -m "Define the device list schema and its row mapping"
```

---

## Task 16: `provisionLists` — idempotent list and column creation

Implements spec §9.1.

**Files:**
- Create: `src/features/devices/sharepoint/spClient.js`
- Create: `src/features/devices/sharepoint/provisionLists.js`
- Test: `src/features/devices/sharepoint/provisionLists.test.js`

**Interfaces:**
- Consumes: `DEVICE_COLUMNS`, `CHANGE_COLUMNS`, the two list names (Task 15).
- Produces:
  - `spFetch(siteUrl, path, { token, digest, method, body, accept }): Promise<Response>`
  - `getFormDigest(siteUrl, token): Promise<string>`
  - `fieldBody(column): object` — the create-field request body for one column
  - `provisionLists(siteUrl, token): Promise<void>`

**Copy `ensureAssetColumns`, not `ensureColumns`.** The older path in
`src/services/sharePointService.js` never puts `col.choices` into the body, so
its choice columns are created with no choices. That bug is out of scope to fix,
but must not be reproduced here.

Two departures from the existing service, both deliberate and both tested:
- DateTime columns get `DisplayFormat: 1` (date **and** time). The existing
  service hardcodes `0` (DateOnly), which discards the time this feature exists
  to record.
- Note columns are created as `SP.FieldMultiLineText` with `RichText: false`.
  A rich-text Note wraps stored values in `<div>` markup and would not
  round-trip.

- [ ] **Step 1: Write the failing test**

`src/features/devices/sharepoint/provisionLists.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { fieldBody } from './provisionLists.js';

describe('fieldBody', () => {
  it('creates a text column as a plain SP.Field', () => {
    expect(fieldBody({ StaticName: 'Owner', Title: 'Owner', kind: 'text' })).toEqual({
      __metadata: { type: 'SP.Field' },
      Title: 'Owner', StaticName: 'Owner', FieldTypeKind: 2, Required: false,
    });
  });

  it('creates a DateTime column with the time kept', () => {
    const body = fieldBody({ StaticName: 'ScannedOn', Title: 'Scanned On', kind: 'datetime' });
    expect(body.__metadata.type).toBe('SP.FieldDateTime');
    expect(body.FieldTypeKind).toBe(4);
    // 1 = DateTime. 0 would be DateOnly and would throw the time away.
    expect(body.DisplayFormat).toBe(1);
  });

  it('creates a Note column as plain text, not rich text', () => {
    const body = fieldBody({ StaticName: 'RawReport', Title: 'Raw Report', kind: 'note' });
    expect(body.__metadata.type).toBe('SP.FieldMultiLineText');
    expect(body.FieldTypeKind).toBe(3);
    expect(body.RichText).toBe(false);
    expect(body.AppendOnly).toBe(false);
  });

  it('creates a choice column WITH its choices', () => {
    const body = fieldBody({
      StaticName: 'DeviceType', Title: 'Device Type', kind: 'choice',
      choices: ['Laptop', 'Desktop', 'Unknown'],
    });
    expect(body.FieldTypeKind).toBe(6);
    expect(body.Choices).toEqual({ results: ['Laptop', 'Desktop', 'Unknown'] });
  });

  it('creates a number column with no decimal places', () => {
    const body = fieldBody({ StaticName: 'InstalledRamGB', Title: 'RAM', kind: 'number' });
    expect(body.__metadata.type).toBe('SP.FieldNumber');
    expect(body.FieldTypeKind).toBe(9);
    expect(body.DisplayFormat).toBe(0);
  });

  it('creates a boolean column', () => {
    const body = fieldBody({ StaticName: 'HasHdd', Title: 'Has HDD', kind: 'boolean' });
    expect(body.FieldTypeKind).toBe(8);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/sharepoint/provisionLists.test.js`
Expected: FAIL — `Failed to resolve import "./provisionLists.js"`

- [ ] **Step 3: Write `spClient.js`**

```js
const VERBOSE = 'application/json;odata=verbose';
const NOMETADATA = 'application/json;odata=nometadata';

export function spFetch(siteUrl, path, { token, digest, method = 'GET', body, accept = VERBOSE }) {
  const headers = {
    Accept: accept,
    'Content-Type': accept,
    Authorization: `Bearer ${token}`,
  };
  if (digest) headers['X-RequestDigest'] = digest;

  return fetch(`${siteUrl}${path}`, {
    method,
    headers,
    body: body === undefined ? undefined : JSON.stringify(body),
  });
}

export async function getFormDigest(siteUrl, token) {
  const response = await spFetch(siteUrl, '/_api/contextinfo', { token, method: 'POST' });
  if (!response.ok) throw new Error(`Could not get a form digest (${response.status})`);

  const data = await response.json();
  const digest = data?.d?.GetContextWebInformation?.FormDigestValue;
  if (!digest) throw new Error('SharePoint returned no form digest');
  return digest;
}

export const ITEM_ACCEPT = NOMETADATA;
export const listPath = (name) => `/_api/web/lists/getByTitle('${encodeURIComponent(name)}')`;
```

- [ ] **Step 4: Write `provisionLists.js`**

```js
import { spFetch, getFormDigest, listPath } from './spClient.js';
import { DEVICE_COLUMNS, CHANGE_COLUMNS, DEVICE_LIST_NAME, CHANGE_LIST_NAME } from './deviceSchema.js';

const FIELD_TYPE_KIND = { text: 2, note: 3, datetime: 4, choice: 6, boolean: 8, number: 9 };

const METADATA_TYPE = {
  text: 'SP.Field',
  choice: 'SP.Field',
  boolean: 'SP.Field',
  note: 'SP.FieldMultiLineText',
  datetime: 'SP.FieldDateTime',
  number: 'SP.FieldNumber',
};

export function fieldBody(column) {
  const body = {
    __metadata: { type: METADATA_TYPE[column.kind] },
    Title: column.Title,
    StaticName: column.StaticName,
    FieldTypeKind: FIELD_TYPE_KIND[column.kind],
    Required: false,
  };

  // DisplayFormat means different things per field type: on DateTime, 1 keeps
  // the time (0 would be date-only); on Number, 0 means zero decimal places.
  if (column.kind === 'datetime') body.DisplayFormat = 1;
  if (column.kind === 'number') body.DisplayFormat = 0;

  if (column.kind === 'note') {
    body.RichText = false;
    body.AppendOnly = false;
    body.NumberOfLines = 6;
  }

  if (column.kind === 'choice') body.Choices = { results: column.choices };

  return body;
}

async function ensureList(siteUrl, token, digest, title, description) {
  const existing = await spFetch(siteUrl, listPath(title), { token });
  if (existing.ok) return;
  if (existing.status !== 404) {
    throw new Error(`Could not check for the "${title}" list (${existing.status})`);
  }

  const created = await spFetch(siteUrl, '/_api/web/lists', {
    token,
    digest,
    method: 'POST',
    body: { __metadata: { type: 'SP.List' }, BaseTemplate: 100, Title: title, Description: description },
  });

  if (!created.ok && created.status !== 201) {
    throw new Error(`Could not create the "${title}" list (${created.status}): ${await created.text()}`);
  }
}

async function existingFieldNames(siteUrl, token, title) {
  const response = await spFetch(siteUrl, `${listPath(title)}/fields?$select=StaticName`, { token });
  if (!response.ok) throw new Error(`Could not read the fields of "${title}" (${response.status})`);

  const data = await response.json();
  return new Set((data.d?.results ?? []).map((field) => field.StaticName));
}

async function ensureColumns(siteUrl, token, digest, title, columns) {
  const existing = await existingFieldNames(siteUrl, token, title);

  for (const column of columns) {
    if (existing.has(column.StaticName)) continue;

    const response = await spFetch(siteUrl, `${listPath(title)}/fields`, {
      token,
      digest,
      method: 'POST',
      body: fieldBody(column),
    });

    // 409 means another tab won the race. The column exists either way.
    if (!response.ok && response.status !== 409) {
      throw new Error(
        `Could not create the "${column.Title}" column (${response.status}): ${await response.text()}`,
      );
    }
  }
}

export async function provisionLists(siteUrl, token) {
  const digest = await getFormDigest(siteUrl, token);

  await ensureList(siteUrl, token, digest, DEVICE_LIST_NAME, 'One row per machine, from the scan reports');
  await ensureColumns(siteUrl, token, digest, DEVICE_LIST_NAME, DEVICE_COLUMNS);

  await ensureList(siteUrl, token, digest, CHANGE_LIST_NAME, 'Field-level change history for the device list');
  await ensureColumns(siteUrl, token, digest, CHANGE_LIST_NAME, CHANGE_COLUMNS);

  return digest;
}
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/sharepoint/provisionLists.test.js`
Expected: PASS — 6 tests.

- [ ] **Step 6: Commit**

```bash
git add src/features/devices/sharepoint/
git commit -m "Provision the device lists and their columns idempotently"
```

**Live verification happens in Task 19**, the first task that actually calls
this against the tenant. If SharePoint rejects `SP.FieldNumber` or
`SP.FieldMultiLineText`, fall back to `SP.Field` with the same `FieldTypeKind`,
keep the test asserting whatever shape works, and record the finding in
`AGENTS.md` (spec §9.1).

---

## Task 17: `diffDevice` and `readDevices`

Implements spec §9.3 (tracked fields), §9.5 (steps 2–4).

**Files:**
- Create: `src/features/devices/sharepoint/diffDevice.js`
- Create: `src/features/devices/sharepoint/readDevices.js`
- Test: `src/features/devices/sharepoint/diffDevice.test.js`

**Interfaces:**
- Consumes: `TRACKED_FIELDS`, `fromListItem` (Task 15), `spFetch`, `listPath` (Task 16).
- Produces:
  - `diffDevice(existing, incoming): { fieldName, oldValue, newValue, changeType }[]`
  - `indexByName(records): Map<string, record>` — keyed on lower-cased computer name
  - `readAllDevices(siteUrl, token): Promise<record[]>` — follows `d.__next` until exhausted

**Paging note:** the reads use `odata=verbose` headers, matching the rest of
this repo, so the continuation link is `data.d.__next` — an absolute URL, not a
path. Fetch it directly rather than re-prefixing `siteUrl`.

- [ ] **Step 1: Write the failing test**

`src/features/devices/sharepoint/diffDevice.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { diffDevice, indexByName } from './diffDevice.js';

const existing = {
  computerName: 'PC1', owner: 'Ali', department: 'SALES', deviceType: 'Laptop',
  computerModel: 'HP 15', windowsVersion: 'Microsoft Windows 10 Pro', osSupported: false,
  cpuModel: 'i5', cpuAgeBand: 'Aging', installedRamGB: 8, ramType: 'DDR4', ramSlotsUsed: 1,
  storageTotalGB: 477, storageType: 'SSD only', antivirusStatus: 'Active', riskLevel: 'High',
  ipAddress: '192.168.1.10', ssid: 'PMW_Group',
};

describe('diffDevice', () => {
  it('finds nothing when nothing changed', () => {
    expect(diffDevice(existing, { ...existing })).toEqual([]);
  });

  it('reports an upgrade as an Updated change', () => {
    const changes = diffDevice(existing, { ...existing, installedRamGB: 16 });
    expect(changes).toEqual([
      { fieldName: 'installedRamGB', oldValue: '8', newValue: '16', changeType: 'Updated' },
    ]);
  });

  it('reports a newly filled field as Added', () => {
    const changes = diffDevice({ ...existing, owner: null }, existing);
    expect(changes).toEqual([
      { fieldName: 'owner', oldValue: '', newValue: 'Ali', changeType: 'Added' },
    ]);
  });

  it('reports a cleared field as Removed', () => {
    const changes = diffDevice(existing, { ...existing, department: null });
    expect(changes).toEqual([
      { fieldName: 'department', oldValue: 'SALES', newValue: '', changeType: 'Removed' },
    ]);
  });

  it('ignores changes to DHCP-volatile fields', () => {
    const changes = diffDevice(existing, { ...existing, ipAddress: '192.168.1.99', ssid: 'Other' });
    expect(changes).toEqual([]);
  });

  it('treats a boolean flip as a change', () => {
    const changes = diffDevice(existing, { ...existing, osSupported: true });
    expect(changes).toEqual([
      { fieldName: 'osSupported', oldValue: 'false', newValue: 'true', changeType: 'Updated' },
    ]);
  });

  it('does not report a change when a number arrives as its string form', () => {
    expect(diffDevice(existing, { ...existing, installedRamGB: '8' })).toEqual([]);
  });

  it('reports several changes at once, in tracked-field order', () => {
    const changes = diffDevice(existing, {
      ...existing, installedRamGB: 16, riskLevel: 'Watch',
    });
    expect(changes.map((c) => c.fieldName)).toEqual(['installedRamGB', 'riskLevel']);
  });
});

describe('indexByName', () => {
  it('keys on a lower-cased computer name', () => {
    const index = indexByName([{ computerName: 'ASHRAF-PC' }]);
    expect(index.get('ashraf-pc')).toBeDefined();
  });

  it('skips rows with no computer name rather than colliding on empty', () => {
    const index = indexByName([{ computerName: null }, { computerName: 'PC1' }]);
    expect(index.size).toBe(1);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/sharepoint/diffDevice.test.js`
Expected: FAIL — `Failed to resolve import "./diffDevice.js"`

- [ ] **Step 3: Write `diffDevice.js`**

```js
import { TRACKED_FIELDS } from './deviceSchema.js';

/**
 * Compared as strings so a number that came back from SharePoint as "8" does
 * not read as a change against the number 8. Null and empty are the same thing
 * here — both mean "we do not have this".
 */
const asText = (value) => (value === null || value === undefined ? '' : String(value));

export function diffDevice(existing, incoming) {
  const changes = [];

  for (const fieldName of TRACKED_FIELDS) {
    const oldValue = asText(existing?.[fieldName]);
    const newValue = asText(incoming?.[fieldName]);

    if (oldValue === newValue) continue;

    let changeType = 'Updated';
    if (!oldValue) changeType = 'Added';
    else if (!newValue) changeType = 'Removed';

    changes.push({ fieldName, oldValue, newValue, changeType });
  }

  return changes;
}

export function indexByName(records) {
  const index = new Map();
  for (const record of records) {
    if (!record.computerName) continue;
    index.set(String(record.computerName).toLowerCase(), record);
  }
  return index;
}
```

- [ ] **Step 4: Write `readDevices.js`**

```js
import { spFetch, listPath } from './spClient.js';
import { DEVICE_LIST_NAME, fromListItem } from './deviceSchema.js';

const PAGE_SIZE = 500;

/**
 * Reads every row once so the whole batch can be diffed in memory. The
 * alternative — one `$filter=Title eq '…'` per dropped file — is one request
 * per file and needs Title indexed to survive the 5,000-item view threshold.
 */
export async function readAllDevices(siteUrl, token) {
  const rows = [];
  let url = `${siteUrl}${listPath(DEVICE_LIST_NAME)}/items?$top=${PAGE_SIZE}`;

  while (url) {
    // `d.__next` is an absolute URL, so after the first page the path is
    // already complete and siteUrl must not be prefixed again.
    const response = await spFetch('', url, { token });

    if (response.status === 404) return [];
    if (!response.ok) {
      throw new Error(`Could not read the device list (${response.status})`);
    }

    const data = await response.json();
    rows.push(...(data.d?.results ?? []));
    url = data.d?.__next ?? null;
  }

  return rows.map(fromListItem);
}
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/sharepoint/diffDevice.test.js`
Expected: PASS — 10 tests.

- [ ] **Step 6: Commit**

```bash
git add src/features/devices/sharepoint/
git commit -m "Diff device rows on tracked fields and read the list in pages"
```

---

## Task 18: `writePool` — concurrency and backoff

Implements spec §9.5 (steps 5–6).

**Files:**
- Create: `src/features/devices/sharepoint/writePool.js`
- Test: `src/features/devices/sharepoint/writePool.test.js`

**Interfaces:**
- Produces:
  - `runPool(items, worker, { concurrency = 4, onProgress }): Promise<{item, value, error}[]>`
  - `withRetry(attempt, { retries = 3, wait = defaultWait }): Promise<Response>`

`attempt` is an async function returning a `Response`. A `429` or `503` is
retried after `Retry-After` seconds when the header is present, otherwise after
an exponentially growing delay. Any other non-OK response is returned as-is for
the caller to turn into a per-row error — the save is not transactional and a
single bad row must not abort the batch.

`wait` is injectable purely so the tests do not actually sleep.

- [ ] **Step 1: Write the failing test**

`src/features/devices/sharepoint/writePool.test.js`:

```js
import { describe, it, expect, vi } from 'vitest';
import { runPool, withRetry } from './writePool.js';

const response = (status, headers = {}) => ({
  ok: status >= 200 && status < 300,
  status,
  headers: { get: (name) => headers[name.toLowerCase()] ?? null },
});

describe('runPool', () => {
  it('returns a result per item, in input order', async () => {
    const results = await runPool([1, 2, 3], async (n) => n * 2, { concurrency: 2 });
    expect(results.map((r) => r.value)).toEqual([2, 4, 6]);
  });

  it('never runs more than `concurrency` workers at once', async () => {
    let active = 0;
    let peak = 0;
    await runPool([1, 2, 3, 4, 5, 6], async () => {
      active += 1;
      peak = Math.max(peak, active);
      await Promise.resolve();
      active -= 1;
    }, { concurrency: 2 });
    expect(peak).toBeLessThanOrEqual(2);
  });

  it('captures a failure without stopping the rest', async () => {
    const results = await runPool([1, 2, 3], async (n) => {
      if (n === 2) throw new Error('nope');
      return n;
    }, { concurrency: 3 });

    expect(results[1].error.message).toBe('nope');
    expect(results[0].value).toBe(1);
    expect(results[2].value).toBe(3);
  });

  it('reports progress as each item finishes', async () => {
    const seen = [];
    await runPool([1, 2, 3], async (n) => n, {
      concurrency: 1,
      onProgress: (done, total) => seen.push(`${done}/${total}`),
    });
    expect(seen).toEqual(['1/3', '2/3', '3/3']);
  });
});

describe('withRetry', () => {
  it('returns a successful response without retrying', async () => {
    const attempt = vi.fn(async () => response(201));
    const result = await withRetry(attempt, { wait: async () => {} });
    expect(result.status).toBe(201);
    expect(attempt).toHaveBeenCalledTimes(1);
  });

  it('retries a 429 and honours Retry-After', async () => {
    const waits = [];
    const attempt = vi.fn()
      .mockResolvedValueOnce(response(429, { 'retry-after': '2' }))
      .mockResolvedValueOnce(response(201));

    const result = await withRetry(attempt, { wait: async (ms) => { waits.push(ms); } });

    expect(result.status).toBe(201);
    expect(waits).toEqual([2000]);
  });

  it('backs off exponentially when there is no Retry-After', async () => {
    const waits = [];
    const attempt = vi.fn(async () => response(503));

    await withRetry(attempt, { retries: 3, wait: async (ms) => { waits.push(ms); } });

    expect(attempt).toHaveBeenCalledTimes(3);
    expect(waits).toEqual([500, 1000]);
  });

  it('does not retry a 400 — a bad row will stay bad', async () => {
    const attempt = vi.fn(async () => response(400));
    const result = await withRetry(attempt, { wait: async () => {} });
    expect(result.status).toBe(400);
    expect(attempt).toHaveBeenCalledTimes(1);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/sharepoint/writePool.test.js`
Expected: FAIL — `Failed to resolve import "./writePool.js"`

- [ ] **Step 3: Write `writePool.js`**

```js
const RETRYABLE = new Set([429, 503]);

const defaultWait = (ms) => new Promise((resolve) => { setTimeout(resolve, ms); });

export async function withRetry(attempt, { retries = 3, wait = defaultWait } = {}) {
  let response;

  for (let tryNumber = 1; tryNumber <= retries; tryNumber += 1) {
    response = await attempt();

    // A 400 means the row is wrong, not that SharePoint is busy. Retrying it
    // just costs the user three times as long to see the same error.
    if (response.ok || !RETRYABLE.has(response.status)) return response;
    if (tryNumber === retries) return response;

    const retryAfter = Number(response.headers?.get?.('Retry-After'));
    await wait(Number.isFinite(retryAfter) && retryAfter > 0
      ? retryAfter * 1000
      : 500 * 2 ** (tryNumber - 1));
  }

  return response;
}

/**
 * Four writes in flight rather than SharePoint's multipart $batch: a hand-built
 * multipart body is easy to get subtly wrong, and this imports 200 machines in
 * well under a minute. $batch stays available as a later optimisation.
 */
export async function runPool(items, worker, { concurrency = 4, onProgress } = {}) {
  const results = new Array(items.length);
  let next = 0;
  let done = 0;

  const runner = async () => {
    while (next < items.length) {
      const index = next;
      next += 1;

      try {
        results[index] = { item: items[index], value: await worker(items[index]), error: null };
      } catch (error) {
        results[index] = { item: items[index], value: null, error };
      }

      done += 1;
      onProgress?.(done, items.length);
    }
  };

  await Promise.all(
    Array.from({ length: Math.min(concurrency, items.length) }, () => runner()),
  );

  return results;
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/sharepoint/writePool.test.js`
Expected: PASS — 9 tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/sharepoint/
git commit -m "Write to SharePoint through a retrying concurrency pool"
```

---

## Task 19: `syncDevices`, the save stage, and live verification

Implements spec §9.5, §10.1 (save stage), §11 (write failures), and the live
check of §9.1 deferred from Task 16.

**Files:**
- Create: `src/features/devices/sharepoint/syncDevices.js`
- Create: `src/features/devices/ui/SaveProgress.jsx`
- Test: `src/features/devices/sharepoint/syncDevices.test.js`
- Modify: `src/pages/DevicesPage.jsx` (the save stage)
- Modify: `src/styles/devices.css` (append the progress styles)

**Interfaces:**
- Consumes: `provisionLists`, `spFetch`, `listPath`, `ITEM_ACCEPT` (Task 16); `readAllDevices`, `diffDevice`, `indexByName` (Task 17); `runPool`, `withRetry` (Task 18); `toListItem`, `TRACKED_FIELDS`, `formatMYT`.
- Produces:
  - `planSync(incoming, existingIndex): { inserts, updates, changeRows }`
  - `syncDevices({ siteUrl, token, devices, changedBy, onProgress }): Promise<{ results, changeCount }>`
  - `<SaveProgress state onRetry onDone />`

`planSync` is split out from `syncDevices` precisely so the decision-making is
testable without a network: everything that decides *what* to write is pure, and
`syncDevices` only performs it.

- [ ] **Step 1: Write the failing test**

`src/features/devices/sharepoint/syncDevices.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { planSync } from './syncDevices.js';
import { indexByName } from './diffDevice.js';

const device = (overrides) => ({
  computerName: 'PC1', owner: 'Ali', department: 'SALES', deviceType: 'Laptop',
  computerModel: 'HP 15', windowsVersion: 'Microsoft Windows 11 Pro', osSupported: true,
  cpuModel: 'i5', cpuAgeBand: 'Current', installedRamGB: 8, ramType: 'DDR4', ramSlotsUsed: 2,
  storageTotalGB: 477, storageType: 'SSD only', antivirusStatus: 'Active', riskLevel: 'Watch',
  scannedOn: Date.UTC(2026, 7, 19, 1, 18), sourceFileName: 'PC1_.txt',
  ...overrides,
});

describe('planSync', () => {
  it('inserts a machine the list has never seen', () => {
    const plan = planSync([device()], indexByName([]));
    expect(plan.inserts).toHaveLength(1);
    expect(plan.updates).toHaveLength(0);
    expect(plan.changeRows).toHaveLength(0);
  });

  it('does nothing for a machine whose tracked fields are unchanged', () => {
    const existing = { ...device(), id: 7 };
    const plan = planSync([device()], indexByName([existing]));
    expect(plan.inserts).toHaveLength(0);
    expect(plan.updates).toHaveLength(0);
  });

  it('updates a machine whose RAM grew, and logs one change row', () => {
    const existing = { ...device(), id: 7 };
    const plan = planSync([device({ installedRamGB: 16 })], indexByName([existing]));

    expect(plan.updates).toHaveLength(1);
    expect(plan.updates[0].id).toBe(7);
    expect(plan.changeRows).toEqual([
      {
        computerName: 'PC1', fieldName: 'installedRamGB',
        oldValue: '8', newValue: '16', changeType: 'Updated',
      },
    ]);
  });

  it('matches an existing machine case-insensitively', () => {
    const existing = { ...device({ computerName: 'pc1' }), id: 7 };
    const plan = planSync([device({ computerName: 'PC1', installedRamGB: 16 })],
      indexByName([existing]));
    expect(plan.updates).toHaveLength(1);
  });

  it('does not update on an untracked change alone', () => {
    const existing = { ...device(), id: 7, ipAddress: '192.168.1.5' };
    const plan = planSync([device({ ipAddress: '192.168.1.99' })], indexByName([existing]));
    expect(plan.updates).toHaveLength(0);
    expect(plan.changeRows).toHaveLength(0);
  });

  it('carries the item body on both inserts and updates', () => {
    const plan = planSync([device()], indexByName([]));
    expect(plan.inserts[0].body.Title).toBe('PC1');
    expect(plan.inserts[0].computerName).toBe('PC1');
  });

  it('counts a new-and-changed batch correctly', () => {
    const existing = { ...device({ computerName: 'PC1' }), id: 7 };
    const plan = planSync(
      [device({ installedRamGB: 16 }), device({ computerName: 'PC2' })],
      indexByName([existing]),
    );
    expect(plan.inserts.map((i) => i.computerName)).toEqual(['PC2']);
    expect(plan.updates.map((u) => u.computerName)).toEqual(['PC1']);
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/sharepoint/syncDevices.test.js`
Expected: FAIL — `Failed to resolve import "./syncDevices.js"`

- [ ] **Step 3: Write `syncDevices.js`**

```js
import { spFetch, listPath, ITEM_ACCEPT } from './spClient.js';
import { provisionLists } from './provisionLists.js';
import { DEVICE_LIST_NAME, CHANGE_LIST_NAME, toListItem } from './deviceSchema.js';
import { readAllDevices } from './readDevices.js';
import { diffDevice, indexByName } from './diffDevice.js';
import { runPool, withRetry } from './writePool.js';
import { formatMYT } from '../../datastudio/time/malaysiaTime.js';

/**
 * Pure: decides what to write. Kept separate from syncDevices so that every
 * insert/update/skip decision is testable without a token or a network.
 */
export function planSync(incoming, existingIndex) {
  const inserts = [];
  const updates = [];
  const changeRows = [];

  for (const device of incoming) {
    const key = String(device.computerName ?? '').toLowerCase();
    const existing = existingIndex.get(key);
    const body = toListItem(device);

    if (!existing) {
      inserts.push({ computerName: device.computerName, body });
      continue;
    }

    const changes = diffDevice(existing, device);
    if (!changes.length) continue;

    updates.push({ computerName: device.computerName, id: existing.id, body });
    for (const change of changes) {
      changeRows.push({ computerName: device.computerName, ...change });
    }
  }

  return { inserts, updates, changeRows };
}

const itemPath = (listName) => `${listPath(listName)}/items`;

export async function syncDevices({ siteUrl, token, devices, changedBy, onProgress }) {
  // Provisioning runs first and throws on failure: a half-created list would
  // fail every row with the same unhelpful message.
  const digest = await provisionLists(siteUrl, token);

  const existing = await readAllDevices(siteUrl, token);
  const plan = planSync(devices, indexByName(existing));

  const post = (path, body, extraHeaders) =>
    withRetry(() =>
      spFetch(siteUrl, path, {
        token, digest, method: 'POST', body, accept: ITEM_ACCEPT, ...extraHeaders,
      }));

  const work = [
    ...plan.inserts.map((entry) => ({ ...entry, action: 'insert' })),
    ...plan.updates.map((entry) => ({ ...entry, action: 'update' })),
  ];

  const results = await runPool(
    work,
    async (entry) => {
      const response = entry.action === 'insert'
        ? await post(itemPath(DEVICE_LIST_NAME), entry.body)
        : await withRetry(() =>
          spFetch(siteUrl, `${itemPath(DEVICE_LIST_NAME)}(${entry.id})`, {
            token,
            digest,
            method: 'POST',
            body: entry.body,
            accept: ITEM_ACCEPT,
            // SharePoint updates are a POST wearing these two headers.
            headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
          }));

      if (!response.ok) {
        throw new Error(`${response.status}: ${await response.text()}`);
      }
      return entry.action;
    },
    { concurrency: 4, onProgress },
  );

  const changedOn = Date.now();
  await runPool(plan.changeRows, async (row) => {
    const response = await post(itemPath(CHANGE_LIST_NAME), {
      Title: row.computerName,
      FieldName: row.fieldName,
      OldValue: row.oldValue,
      NewValue: row.newValue,
      ChangeType: row.changeType,
      ChangedOn: new Date(changedOn).toISOString(),
      ChangedOnMYT: formatMYT(changedOn, 'datetime12'),
      ChangedBy: changedBy ?? '',
    });
    if (!response.ok) throw new Error(`${response.status}`);
  }, { concurrency: 4 });

  return {
    results: results.map((result, index) => ({
      computerName: work[index].computerName,
      action: work[index].action,
      error: result.error ? result.error.message : null,
    })),
    changeCount: plan.changeRows.length,
    skipped: devices.length - work.length,
  };
}
```

**On the update request:** SharePoint's REST update is a `POST` carrying
`X-HTTP-Method: MERGE` and `IF-MATCH: *`. That means `spFetch` (Task 16) needs a
`headers` option merged over the ones it builds — add it there:

```js
export function spFetch(siteUrl, path, { token, digest, method = 'GET', body, accept = VERBOSE, headers: extra }) {
  const headers = { Accept: accept, 'Content-Type': accept, Authorization: `Bearer ${token}`, ...extra };
```

with `digest` still applied after the spread so a caller cannot accidentally
drop it. The existing service does the equivalent with `PATCH` in
`updateListItem`; either works, and Step 8 proves whichever you pick.

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/sharepoint/syncDevices.test.js`
Expected: PASS — 7 tests.

- [ ] **Step 5: Write `SaveProgress.jsx`**

```jsx
import Button from '../../../components/ui/Button';
import { Check, AlertTriangle } from '../../../components/ui/Icons';

export default function SaveProgress({ state, onRetry, onDone }) {
  const { done, total, results, error } = state;
  const failures = (results ?? []).filter((row) => row.error);
  const finished = results !== null;

  if (error) {
    return (
      <div className="sp-status">
        <AlertTriangle size={18} />
        <p>{error}</p>
        <Button variant="secondary" onClick={onRetry}>Try again</Button>
      </div>
    );
  }

  if (!finished) {
    return (
      <div className="sp-status">
        <div className="sp-bar" role="progressbar" aria-valuenow={done} aria-valuemax={total}>
          <span style={{ width: `${total ? (done / total) * 100 : 0}%` }} />
        </div>
        <p>Saving {done} of {total}…</p>
      </div>
    );
  }

  return (
    <div className="sp-status">
      {failures.length === 0 ? <Check size={18} /> : <AlertTriangle size={18} />}
      <p>
        {results.length - failures.length} saved
        {failures.length > 0 && `, ${failures.length} failed`}
      </p>

      {failures.length > 0 && (
        <ul className="sp-failures">
          {failures.map((row) => (
            <li key={row.computerName}>
              <strong>{row.computerName}</strong> — {row.error}
            </li>
          ))}
        </ul>
      )}

      <div className="sp-actions">
        {failures.length > 0 && (
          <Button variant="secondary" onClick={() => onRetry(failures.map((f) => f.computerName))}>
            Retry the {failures.length} that failed
          </Button>
        )}
        <Button onClick={onDone}>Done</Button>
      </div>
    </div>
  );
}
```

- [ ] **Step 6: Wire the save stage into `DevicesPage.jsx`**

Use the repo's token hook — never `acquireTokenSilent` directly:

```jsx
import { useSharePointToken } from '../hooks/useRequests';
import { syncDevices } from '../features/devices/sharepoint/syncDevices';
import SaveProgress from '../features/devices/ui/SaveProgress';

const SITE_URL = import.meta.env.VITE_SHAREPOINT_SITE_URL;
```

and, inside the component:

```jsx
  const getToken = useSharePointToken();
  const [save, setSave] = useState({ done: 0, total: 0, results: null, error: null });

  const handleSave = async (only) => {
    const toSave = merged.filter(
      (device) =>
        !excluded.has(device.sourceFileName) &&
        (!only || only.includes(device.computerName)),
    );

    setStage('save');
    setSave({ done: 0, total: toSave.length, results: null, error: null });

    try {
      const token = await getToken();
      const outcome = await syncDevices({
        siteUrl: SITE_URL,
        token,
        devices: toSave,
        changedBy: account?.username ?? '',
        onProgress: (done, total) => setSave((s) => ({ ...s, done, total })),
      });
      setSave((s) => ({ ...s, results: outcome.results }));
    } catch (error) {
      setSave((s) => ({ ...s, error: error.message }));
    }
  };
```

The review stage gets a Save button that calls `handleSave()`, and the save
stage renders `<SaveProgress state={save} onRetry={handleSave} onDone={reset} />`.

Confirm the exact export name and signature of the token hook in
`src/hooks/useRequests.js` before wiring — match it, do not invent it.

- [ ] **Step 7: Append the progress styles to `devices.css`**

```css
.sp-status {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 10px;
  padding: 28px 20px;
  text-align: center;
  color: var(--it-ink);
}

.sp-bar {
  width: min(420px, 100%);
  height: 6px;
  border-radius: 999px;
  background: var(--it-brand-wash);
  overflow: hidden;
}

.sp-bar span {
  display: block;
  height: 100%;
  background: var(--it-brand);
  transition: width 160ms ease;
}

.sp-failures {
  margin: 0;
  padding-left: 18px;
  text-align: left;
  font-size: 0.82rem;
  color: var(--it-danger);
}

.sp-actions { display: flex; gap: 8px; }

@media (prefers-reduced-motion: reduce) {
  .sp-bar span { transition: none; }
}
```

- [ ] **Step 8: Verify against the live tenant**

This is the first task that touches SharePoint, and it is where the two
provisioning departures from the existing service get proven:

1. Drop three real files and save.
2. Open the SharePoint site. Confirm `IT Device List` exists with the columns
   from Task 15, and specifically that:
   - `Scanned On` shows a **time as well as a date** (proves `DisplayFormat: 1`).
   - `Raw Report` contains plain text with **no `<div>` markup** (proves
     `RichText: false`).
   - `Device Type` offers its three choices in the column's dropdown.
   - `Installed RAM (GB)` sorts numerically, not as text.
3. Compare `Scanned On (MYT)` against `Scanned On`. If they disagree by a whole
   number of hours, the site's regional timezone is not UTC+8 — that is the
   case the mirror column exists for, and spec §9.4 says to surface it.
4. Save the same three files again unchanged. Confirm **no** new rows and **no**
   change-log rows.
5. Edit one owner in the review grid, save again. Confirm the row updates and
   exactly one `IT Device Changes` row appears.
6. If `SP.FieldNumber` or `SP.FieldMultiLineText` was rejected, apply the
   documented fallback (`SP.Field` + the same `FieldTypeKind`), update the Task
   16 test to assert what actually works, and note it in `AGENTS.md`.

- [ ] **Step 9: Commit**

```bash
git add src/ && git commit -m "Sync reviewed devices to SharePoint with a change log"
```

---

# PHASE 5 — The register

---

## Task 20: `useDevices` and the register table

Implements spec §10.2.

**Files:**
- Create: `src/features/devices/useDevices.js`
- Create: `src/features/devices/ui/DeviceTable.jsx`
- Create: `src/features/devices/deviceFilters.js`
- Test: `src/features/devices/deviceFilters.test.js`
- Modify: `src/pages/DevicesPage.jsx` (a register tab beside the import flow)
- Modify: `src/styles/devices.css`

**Interfaces:**
- Consumes: `readAllDevices` (Task 17), `useSharePointToken` from `src/hooks/useRequests.js`, `formatMYT`.
- Produces:
  - `applyFilters(devices, params): device[]` — `params` is a plain object of the query-string values
  - `toCsv(devices, columns): string`
  - `useDevices(): { devices, loading, error, reload }`
  - `<DeviceTable devices onFilter />`

**Filter keys**, matching the query string in spec §10.2: `risk`, `type`,
`department`, `os`, `storage`, `ram`, `q`. `ram` is a bucket label such as
`8 GB`; `os` is `Supported` or `Unsupported`; `q` searches computer name and
owner.

Follow the repo convention: the dashboard links into the register with these
keys, and both read the same `useDevices()` fetch, so a figure and the rows
behind it cannot disagree.

- [ ] **Step 1: Write the failing test**

`src/features/devices/deviceFilters.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { applyFilters, toCsv, ramBucket } from './deviceFilters.js';

const rows = [
  { computerName: 'A', owner: 'Ali', riskLevel: 'Critical', deviceType: 'Desktop',
    department: 'SALES', osSupported: false, storageType: 'Mixed', installedRamGB: 2 },
  { computerName: 'B', owner: 'Bea', riskLevel: 'OK', deviceType: 'Laptop',
    department: 'FINANCE', osSupported: true, storageType: 'SSD only', installedRamGB: 16 },
  { computerName: 'C', owner: null, riskLevel: 'Watch', deviceType: 'Laptop',
    department: 'SALES', osSupported: true, storageType: 'SSD only', installedRamGB: 8 },
];

describe('ramBucket', () => {
  it('buckets by installed size', () => {
    expect(ramBucket(2)).toBe('2 GB');
    expect(ramBucket(16)).toBe('16 GB');
    expect(ramBucket(null)).toBe('Unknown');
  });
});

describe('applyFilters', () => {
  it('returns everything when no filter is set', () => {
    expect(applyFilters(rows, {})).toHaveLength(3);
  });

  it('filters by risk level', () => {
    expect(applyFilters(rows, { risk: 'Critical' }).map((r) => r.computerName)).toEqual(['A']);
  });

  it('filters by device type', () => {
    expect(applyFilters(rows, { type: 'Laptop' }).map((r) => r.computerName)).toEqual(['B', 'C']);
  });

  it('filters by department', () => {
    expect(applyFilters(rows, { department: 'SALES' })).toHaveLength(2);
  });

  it('filters unsupported operating systems', () => {
    expect(applyFilters(rows, { os: 'Unsupported' }).map((r) => r.computerName)).toEqual(['A']);
  });

  it('filters by storage type and RAM bucket', () => {
    expect(applyFilters(rows, { storage: 'SSD only' })).toHaveLength(2);
    expect(applyFilters(rows, { ram: '8 GB' }).map((r) => r.computerName)).toEqual(['C']);
  });

  it('searches computer name and owner, case-insensitively', () => {
    expect(applyFilters(rows, { q: 'bea' }).map((r) => r.computerName)).toEqual(['B']);
    expect(applyFilters(rows, { q: 'a' }).map((r) => r.computerName)).toEqual(['A']);
  });

  it('combines filters', () => {
    expect(applyFilters(rows, { department: 'SALES', risk: 'Watch' }).map((r) => r.computerName))
      .toEqual(['C']);
  });
});

describe('toCsv', () => {
  it('writes a header row and quotes what needs quoting', () => {
    const csv = toCsv(
      [{ computerName: 'A, Ltd', owner: 'He said "hi"' }],
      [{ key: 'computerName', label: 'Computer' }, { key: 'owner', label: 'Owner' }],
    );
    expect(csv).toBe('Computer,Owner\r\n"A, Ltd","He said ""hi"""');
  });

  it('renders a null as an empty cell', () => {
    expect(toCsv([{ owner: null }], [{ key: 'owner', label: 'Owner' }])).toBe('Owner\r\n');
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/deviceFilters.test.js`
Expected: FAIL — `Failed to resolve import "./deviceFilters.js"`

- [ ] **Step 3: Write `deviceFilters.js`**

```js
export function ramBucket(installedRamGB) {
  return typeof installedRamGB === 'number' ? `${installedRamGB} GB` : 'Unknown';
}

const MATCHERS = {
  risk: (device, value) => device.riskLevel === value,
  type: (device, value) => device.deviceType === value,
  department: (device, value) => device.department === value,
  storage: (device, value) => device.storageType === value,
  ram: (device, value) => ramBucket(device.installedRamGB) === value,
  os: (device, value) =>
    value === 'Unsupported' ? device.osSupported === false : device.osSupported === true,
  q: (device, value) => {
    const needle = value.toLowerCase();
    return `${device.computerName ?? ''} ${device.owner ?? ''}`.toLowerCase().includes(needle);
  },
};

export function applyFilters(devices, params) {
  return devices.filter((device) =>
    Object.entries(params).every(([key, value]) => {
      if (!value) return true;
      const matcher = MATCHERS[key];
      return matcher ? matcher(device, value) : true;
    }));
}

const cell = (value) => {
  if (value === null || value === undefined) return '';
  const text = Array.isArray(value) ? value.join('; ') : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
};

export function toCsv(devices, columns) {
  const header = columns.map((column) => cell(column.label)).join(',');
  const body = devices.map((device) =>
    columns.map((column) => cell(device[column.key])).join(','));
  return [header, ...body].join('\r\n');
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/deviceFilters.test.js`
Expected: PASS — 11 tests.

- [ ] **Step 5: Write `useDevices.js`**

```js
import { useCallback, useEffect, useState } from 'react';
import { useSharePointToken } from '../../hooks/useRequests';
import { readAllDevices } from './sharepoint/readDevices';

const SITE_URL = import.meta.env.VITE_SHAREPOINT_SITE_URL;

/** The one SharePoint read for this section. Register and dashboard share it. */
export function useDevices() {
  const getToken = useSharePointToken();
  const [devices, setDevices] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const token = await getToken();
      setDevices(await readAllDevices(SITE_URL, token));
    } catch (failure) {
      setError(failure.message);
    } finally {
      setLoading(false);
    }
  }, [getToken]);

  useEffect(() => { load(); }, [load]);

  return { devices, loading, error, reload: load };
}
```

Confirm `useSharePointToken`'s exact export name and return shape in
`src/hooks/useRequests.js` and match it. Do not call MSAL directly (Global
Constraints).

- [ ] **Step 6: Write `DeviceTable.jsx`**

A sortable table over `applyFilters`, with a search box, the active filters
shown as removable chips, and a Download CSV button that builds a Blob from
`toCsv` and clicks an object URL. Columns: Computer, Owner, Department, Type,
Model, CPU, RAM, Storage, Windows, Antivirus, Risk, Scanned On — the last
rendered with `formatMYT(device.scannedOn, 'datetime12')`.

Read the filters from the query string with `useSearchParams` from
`react-router-dom`, and write them back on change, so a filtered view is a
shareable URL and the dashboard's click-through lands correctly.

- [ ] **Step 7: Verify in the browser**

Load `/devices?risk=Critical`. Confirm only the critical machines appear, the
chip shows, removing it restores the full list, and the CSV downloads with the
filtered rows only.

- [ ] **Step 8: Commit**

```bash
git add src/ && git commit -m "Add the device register with query-string filters and CSV export"
```

---

# PHASE 6 — The dashboard

---

## Task 21: `deviceStats` — the aggregations

Implements spec §10.3, §8.6 (incomplete scans excluded).

**Files:**
- Create: `src/features/devices/stats/deviceStats.js`
- Test: `src/features/devices/stats/deviceStats.test.js`

**Interfaces:**
- Consumes: `ramBucket` (Task 20).
- Produces:
  - `fleetSummary(devices, now): { total, needsAttention, unsupportedOs, unprotected, avgRamGB, staleScans }`
  - `countBy(devices, keyFn): { label, count }[]` — sorted by count descending
  - `scansByMonth(devices): { label, count }[]` — chronological
  - `leaderboards(devices, now): { highestRam, lowestRam, oldest, recent, upgradeCandidates, rescanNeeded }`

**The exclusion rule is the point of this module.** Every figure except
`rescanNeeded` counts only `scanComplete !== false` machines. One failed scan
otherwise drags the fleet average down and shows as a healthy machine.

- [ ] **Step 1: Write the failing test**

`src/features/devices/stats/deviceStats.test.js`:

```js
import { describe, it, expect } from 'vitest';
import { fleetSummary, countBy, scansByMonth, leaderboards } from './deviceStats.js';

const NOW = Date.UTC(2026, 7, 21);
const DAY = 86_400_000;

const rows = [
  { computerName: 'CRIT', riskLevel: 'Critical', osSupported: false, avProtected: true,
    installedRamGB: 2, ramUpgradable: false, cpuAgeBand: 'Obsolete', deviceType: 'Desktop',
    scanComplete: true, scannedOn: NOW - 10 * DAY },
  { computerName: 'OLD', riskLevel: 'High', osSupported: false, avProtected: false,
    installedRamGB: 8, ramUpgradable: true, cpuAgeBand: 'Aging', deviceType: 'Laptop',
    scanComplete: true, scannedOn: NOW - 200 * DAY },
  { computerName: 'GOOD', riskLevel: 'OK', osSupported: true, avProtected: true,
    installedRamGB: 16, ramUpgradable: false, cpuAgeBand: 'Current', deviceType: 'Laptop',
    scanComplete: true, scannedOn: NOW - DAY },
  { computerName: 'BROKEN', riskLevel: 'Unknown', osSupported: null, avProtected: false,
    installedRamGB: null, ramUpgradable: false, cpuAgeBand: 'Unknown', deviceType: 'Unknown',
    scanComplete: false, scannedOn: NOW - 2 * DAY },
];

describe('fleetSummary', () => {
  const summary = fleetSummary(rows, NOW);

  it('counts only complete scans', () => {
    expect(summary.total).toBe(3);
  });

  it('counts Critical and High together as needing attention', () => {
    expect(summary.needsAttention).toBe(2);
  });

  it('counts unsupported operating systems', () => {
    expect(summary.unsupportedOs).toBe(2);
  });

  it('counts unprotected machines', () => {
    expect(summary.unprotected).toBe(1);
  });

  it('averages RAM over complete scans only', () => {
    // (2 + 8 + 16) / 3 — the failed scan must not pull it down
    expect(summary.avgRamGB).toBe(9);
  });

  it('counts scans older than 180 days as stale', () => {
    expect(summary.staleScans).toBe(1);
  });
});

describe('countBy', () => {
  it('counts by a key, biggest group first, excluding incomplete scans', () => {
    expect(countBy(rows, (d) => d.deviceType)).toEqual([
      { label: 'Laptop', count: 2 },
      { label: 'Desktop', count: 1 },
    ]);
  });

  it('labels a missing value rather than dropping the row', () => {
    expect(countBy([{ scanComplete: true, department: null }], (d) => d.department))
      .toEqual([{ label: 'Unassigned', count: 1 }]);
  });
});

describe('scansByMonth', () => {
  it('groups by month in chronological order', () => {
    const result = scansByMonth([
      { scanComplete: true, scannedOn: Date.UTC(2026, 6, 3) },
      { scanComplete: true, scannedOn: Date.UTC(2026, 7, 1) },
      { scanComplete: true, scannedOn: Date.UTC(2026, 7, 20) },
    ]);
    expect(result).toEqual([
      { label: '07/2026', count: 1 },
      { label: '08/2026', count: 2 },
    ]);
  });
});

describe('leaderboards', () => {
  const boards = leaderboards(rows, NOW);

  it('ranks the most and least RAM', () => {
    expect(boards.highestRam[0].computerName).toBe('GOOD');
    expect(boards.lowestRam[0].computerName).toBe('CRIT');
  });

  it('ranks the oldest hardware by CPU age band', () => {
    expect(boards.oldest[0].computerName).toBe('CRIT');
  });

  it('lists the newest scans first', () => {
    expect(boards.recent[0].computerName).toBe('GOOD');
  });

  it('lists only machines fixable with a stick as upgrade candidates', () => {
    expect(boards.upgradeCandidates.map((d) => d.computerName)).toEqual(['OLD']);
  });

  it('lists incomplete and stale scans as needing a re-scan', () => {
    expect(boards.rescanNeeded.map((d) => d.computerName).sort()).toEqual(['BROKEN', 'OLD']);
  });

  it('keeps the failed scan out of every other board', () => {
    for (const board of ['highestRam', 'lowestRam', 'oldest', 'recent', 'upgradeCandidates']) {
      expect(boards[board].some((d) => d.computerName === 'BROKEN')).toBe(false);
    }
  });
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `npx vitest run src/features/devices/stats/deviceStats.test.js`
Expected: FAIL — `Failed to resolve import "./deviceStats.js"`

- [ ] **Step 3: Write `deviceStats.js`**

```js
const STALE_DAYS = 180;
const STALE_MS = STALE_DAYS * 86_400_000;

/**
 * A scan that failed is not a healthy machine — it is an unknown one. Counting
 * it would pull the fleet's average RAM down and park a machine with no CPU at
 * the top of the "all clear" list, so every figure but rescanNeeded ignores it.
 */
const complete = (devices) => devices.filter((device) => device.scanComplete !== false);

const isStale = (device, now) =>
  typeof device.scannedOn === 'number' && now - device.scannedOn > STALE_MS;

export function fleetSummary(devices, now = Date.now()) {
  const rows = complete(devices);
  const ram = rows.map((d) => d.installedRamGB).filter((n) => typeof n === 'number');

  return {
    total: rows.length,
    needsAttention: rows.filter((d) => d.riskLevel === 'Critical' || d.riskLevel === 'High').length,
    unsupportedOs: rows.filter((d) => d.osSupported === false).length,
    unprotected: rows.filter((d) => d.avProtected === false).length,
    avgRamGB: ram.length ? Math.round(ram.reduce((a, b) => a + b, 0) / ram.length) : null,
    staleScans: rows.filter((d) => isStale(d, now)).length,
  };
}

export function countBy(devices, keyFn) {
  const counts = new Map();

  for (const device of complete(devices)) {
    const label = keyFn(device) ?? 'Unassigned';
    counts.set(label, (counts.get(label) ?? 0) + 1);
  }

  return [...counts]
    .map(([label, count]) => ({ label, count }))
    .sort((a, b) => b.count - a.count || a.label.localeCompare(b.label));
}

export function scansByMonth(devices) {
  const counts = new Map();

  for (const device of complete(devices)) {
    if (typeof device.scannedOn !== 'number') continue;
    const date = new Date(device.scannedOn);
    const key = `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
    counts.set(key, (counts.get(key) ?? 0) + 1);
  }

  return [...counts]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([key, count]) => {
      const [year, month] = key.split('-');
      return { label: `${month}/${year}`, count };
    });
}

const AGE_RANK = { Obsolete: 0, Aging: 1, Unknown: 2, Current: 3 };

export function leaderboards(devices, now = Date.now()) {
  const rows = complete(devices);
  const withRam = rows.filter((d) => typeof d.installedRamGB === 'number');

  return {
    highestRam: [...withRam].sort((a, b) => b.installedRamGB - a.installedRamGB).slice(0, 5),
    lowestRam: [...withRam].sort((a, b) => a.installedRamGB - b.installedRamGB).slice(0, 5),
    oldest: [...rows]
      .sort((a, b) => AGE_RANK[a.cpuAgeBand] - AGE_RANK[b.cpuAgeBand])
      .slice(0, 5),
    recent: [...rows].sort((a, b) => (b.scannedOn ?? 0) - (a.scannedOn ?? 0)).slice(0, 5),
    // The cheap fix: a free slot means one stick, not a new machine.
    upgradeCandidates: rows.filter(
      (d) => d.ramUpgradable && typeof d.installedRamGB === 'number' && d.installedRamGB <= 8,
    ),
    rescanNeeded: devices.filter((d) => d.scanComplete === false || isStale(d, now)),
  };
}
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `npx vitest run src/features/devices/stats/deviceStats.test.js`
Expected: PASS — 17 tests.

- [ ] **Step 5: Commit**

```bash
git add src/features/devices/stats/
git commit -m "Aggregate fleet statistics, excluding failed scans"
```

---

## Task 22: The dashboard

Implements spec §10.3.

**Files:**
- Create: `src/features/devices/ui/DeviceCharts.jsx`
- Create: `src/features/devices/ui/Leaderboards.jsx`
- Modify: `src/pages/DevicesPage.jsx` (a dashboard tab)
- Modify: `src/styles/devices.css`

**Interfaces:**
- Consumes: `fleetSummary`, `countBy`, `scansByMonth`, `leaderboards` (Task 21); `useDevices` (Task 20); `StatCard` from `src/components/ui/StatCard.jsx`; `Card` from `Surfaces.jsx`.

**Reuse, do not rebuild.** `src/pages/DashboardPage.jsx` already contains a
`BarChart` and a `ColumnChart` built from CSS. Read them first and match their
markup and class names so the two dashboards look like one product. If a chart
there is directly reusable, lift it into `src/components/ui/` and import it from
both — but do not export it from a file that also default-exports a page
component (Global Constraints).

- [ ] **Step 1: Six stat cards**

Using `StatCard`, which takes `{ icon, label, value, unit, color, loading, onClick }`
and becomes a button when `onClick` is passed:

| Card | Value | Click target |
|---|---|---|
| Total devices | `summary.total` | `/devices?tab=register` |
| Needs attention | `summary.needsAttention` | `?risk=Critical` |
| Unsupported OS | `summary.unsupportedOs` | `?os=Unsupported` |
| Unprotected | `summary.unprotected` | `?av=Unprotected` |
| Average RAM | `summary.avgRamGB` unit `GB` | no drill-down |
| Stale scans | `summary.staleScans` | `?stale=1` |

Add `av` and `stale` to `MATCHERS` in `deviceFilters.js`, with a test each, so
these two cards drill through like the others rather than being dead ends:

```js
  av: (device, value) =>
    value === 'Unprotected' ? device.avProtected === false : device.avProtected === true,
  stale: (device, value, now = Date.now()) =>
    value !== '1' || (typeof device.scannedOn === 'number' && now - device.scannedOn > 180 * 86_400_000),
```

- [ ] **Step 2: Eight charts in `DeviceCharts.jsx`**

Each is `countBy` piped into a bar chart, and each bar links into the register
with its own filter:

| Chart | Data | Filter key |
|---|---|---|
| Risk mix | `countBy(devices, d => d.riskLevel)` | `risk` |
| RAM distribution | `countBy(devices, d => ramBucket(d.installedRamGB))` | `ram` |
| Laptop vs Desktop | `countBy(devices, d => d.deviceType)` | `type` |
| Windows mix | `countBy(devices, d => d.windowsVersion)` | — |
| Storage type | `countBy(devices, d => d.storageType)` | `storage` |
| CPU age | `countBy(devices, d => d.cpuAgeBand)` | — |
| By department | `countBy(devices, d => d.department)` | `department` |
| Scans per month | `scansByMonth(devices)` | — |

Colour rules, tokens only: `Critical`/`High` and any Windows 10 row use
`--it-danger`; `Watch` uses `--it-accent`; `OK` uses `--it-good`; everything
else uses `--it-brand`.

- [ ] **Step 3: Six leaderboards in `Leaderboards.jsx`**

Each is a small `Card` with a title, a one-line explanation, and up to five
rows of `computer — owner — figure`. `Upgrade candidates` shows
`8 GB in 1 of 2 slots`, which is the whole point of the board. `Re-scan needed`
shows the reason: `Scan incomplete` or `Last seen 14/02/2026 10:02 AM`.

- [ ] **Step 4: Verify in the browser**

With real data saved, confirm: the six cards match hand-counted figures, every
bar and card click lands on a register view whose row count equals the figure
clicked, and `CARMEN-HP` appears **only** under Re-scan needed.

Then `resize_window` to mobile and confirm the grid stacks to one column and the
table scrolls inside its own container rather than the page scrolling sideways.

- [ ] **Step 5: Commit**

```bash
git add src/ && git commit -m "Add the device fleet dashboard"
```

---

## Task 23: Document the section

**Files:**
- Modify: `AGENTS.md`

- [ ] **Step 1: Add the route and the locations**

Add `/devices` to the ROUTES table, and to WHERE TO LOOK:

| Task | Location |
|---|---|
| Device report parsing | `src/features/devices/parse/` |
| Device derived fields and risk | `src/features/devices/derive/` |
| Device SharePoint schema | `src/features/devices/sharepoint/deviceSchema.js` |
| Device fleet statistics | `src/features/devices/stats/deviceStats.js` |

- [ ] **Step 2: Add the conventions this feature establishes**

Under CONVENTIONS:

- **`src/features/<name>/`** is where a section with more than a handful of
  modules lives. `datastudio/` and `devices/` follow it. Layering inside a
  feature: `parse/` knows nothing about the domain, `derive/` knows nothing
  about SharePoint, `sharepoint/` imports no React.
- **Device report parsing keys off a known-label whitelist.** A generic
  `^Word:` split reads `Total Slots: 2 | Used Slots: 2` and `Y: | \\server\PMW`
  as field names and moves those values out of the blocks they belong to.
- **`Total RAM` in a scan report is usable RAM, not installed RAM.** Sum
  `RAM Slot Info` for the real figure.

Under ANTI-PATTERNS:

- Don't create a SharePoint DateTime column with `DisplayFormat: 0` when the
  time matters — that is DateOnly and silently discards it.
- Don't create a Note column without `RichText: false`; a rich-text Note stores
  `<div>` markup around the value and will not round-trip.
- Don't add `hour12` beside `hourCycle` in `malaysiaTime.js` — an explicit
  `hour12` nullifies `hourCycle` entirely.

- [ ] **Step 3: Record whatever the tenant actually accepted**

If Task 19's live check forced a fallback from `SP.FieldNumber` or
`SP.FieldMultiLineText`, write down what worked and what did not, so the next
person does not rediscover it.

- [ ] **Step 4: Full verification**

```bash
npm test && npm run lint && npm run build
```

Expected: all tests pass; no new lint errors (the four pre-existing failures
remain); the build produces a `dist/index.html` containing a real bundle
reference, not an `export default` string.

- [ ] **Step 5: Commit**

```bash
git add AGENTS.md && git commit -m "Document the device list section"
```

---

## Plan self-review

**Spec coverage.** §4.1 → Task 13 (separate route). §4.2 → Tasks 10, 15 (both
dates). §4.3 → Tasks 17, 19 (upsert plus change log). §4.4 → Tasks 10, 15 (raw
alongside derived). §4.5 → Task 22 (no chart library). §4.6 → Tasks 12–14 (parse
before token). §7 → Tasks 1–3. §8 → Tasks 4–10. §9.1–§9.3 → Tasks 15, 16.
§9.4 → Tasks 11, 15, 19. §9.5 → Tasks 17–19. §10.1 → Tasks 13, 14, 19.
§10.2 → Task 20. §10.3 → Tasks 21, 22. §11 → Tasks 12 (rejections), 19 (write
failures). §12 → every task's tests. §13 → the phase headings. §14 → Tasks 16,
19 (the tenant fallback).

**One spec item is deliberately deferred, not dropped:** the site-timezone
notice (§9.4, last paragraph) is *checked* in Task 19 Step 8 but only *rendered*
if the check shows a mismatch. If it does, add the notice to `DevicesPage`
before closing Task 19 — a one-line banner reading which timezone the site is
set to and pointing at the mirror column.

**Type consistency.** `DeviceRecord` keys are fixed in Task 10 and used
unchanged in Tasks 12, 14, 15, 17, 20, 21. `StaticName` → camelCase mapping is
defined once, in `keyFor` (Task 15). `issuesFor` / `sortForReview` (Task 14),
`applyFilters` / `ramBucket` / `toCsv` (Task 20), and `fleetSummary` / `countBy`
/ `scansByMonth` / `leaderboards` (Task 21) each appear under one name only.
`spFetch` gains its `headers` option in Task 16 and is used with it in Task 19.

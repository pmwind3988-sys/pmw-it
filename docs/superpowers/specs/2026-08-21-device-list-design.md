# Device List — .txt scan import, SharePoint register and fleet dashboard

**Date:** 2026-08-21
**Status:** Approved design, ready for implementation planning
**Route:** `/devices`
**Branch:** `feat/data-studio`

## 1. Purpose

IT collects a plain-text spec report from each machine in the company. Today those
files pile up in a folder and nobody can answer "which machines need attention?"
without opening them one by one.

This section takes a drop of those `.txt` files, parses each into a row, derives the
numbers that make the fleet sortable, lets the user correct what was guessed, and
writes the result to a SharePoint list that is useful both inside this portal and to
anyone who opens the list directly. A dashboard on top answers the fleet questions:
what needs attention, what is old, what is new, who has the most and least RAM, how
much storage, laptop or desktop, and where the risk is.

## 2. Goals

1. Drag-and-drop (or file-pick) many `.txt` reports at once.
2. Parse each file into fields, correctly separating one answer from many.
3. Derive typed, sortable columns from the answers — RAM in GB, storage in GB, CPU
   generation, device type, risk score.
4. Show a review table before anything is written, with guessed values flagged and
   editable.
5. Store one row per machine in SharePoint, with accurate column types and Malaysian
   time including AM/PM.
6. Re-import updates the existing row and records what changed.
7. A dashboard of statistics and leaderboards, every figure clicking through to the
   rows behind it.

## 3. Non-goals (v1)

| Excluded | Reason |
|---|---|
| Editing rows after they are saved | The scan file is the source of truth; correct it at import |
| Deleting devices from the portal | Deletion is a SharePoint permission question, not a UI feature |
| Agent / remote collection | The existing scan script already produces the files |
| Per-user permissions inside the section | `AppShell` is the one auth gate; SharePoint enforces the rest |
| Historical spec charts over time | The change log records history; charting it is its own project |
| Excel or CSV import | That is Data Studio's job (`2026-08-20-data-studio-design.md`) |
| Free-disk-space reporting | The scan file does not contain it |

## 4. Decisions

### 4.1 A separate section, not part of Data Studio

Data Studio is a generic BI tool over arbitrary spreadsheets, stores nothing outside
the browser, and infers its schema per import. This is a fixed-schema register backed
by SharePoint with a domain-specific dashboard. They share only date handling.

**Consequence:** `src/features/devices/` is a sibling of `src/features/datastudio/`.
The only shared import is `../datastudio/time/malaysiaTime.js`. Nothing in Data Studio
imports from `devices/`.

### 4.2 Two dates, stored separately

`Scanned On` comes from the browser's `File.lastModified` — when the scan was written.
`Imported On` is the moment of the save. Keeping both means staleness ("this machine
has not been scanned in 14 months") and audit ("this row entered the register today")
are separately answerable. `Scanned On` is editable in review, because copying files
between machines can reset a file timestamp.

### 4.3 Upsert by Computer Name, with a change log

One row per machine, so the list always reads as current fleet state. A second list
records each changed field. Volatile fields are excluded from the log (§9.3).

### 4.4 Raw answers and derived columns side by side

Every field keeps its verbatim text. Derived columns sit beside it. Without the derived
columns "who has the least RAM" cannot be sorted, because `"8 GB"` is a string; without
the raw text, information is lost the first time a derivation rule is wrong.

### 4.5 No chart library

`DashboardPage.jsx` already renders bar and column charts from CSS. The device
dashboard reuses that approach. Adding ECharts for eight bar charts would cost ~200KB
gzipped for no visual gain.

### 4.6 Parse first, authenticate later

Parsing is pure and runs on drop with no token. The user sees the full review table
before any SharePoint call happens. A token is acquired only at save.

## 5. Pipeline

```
File[] → parseReport   (pure)  → { fields, unknownLabels, warnings }
       → deriveDevice  (pure)  → typed device record + risk
       → ReviewGrid           → user corrections, exclusions
       → syncDevices          → read all → diff → insert/update + change rows
       → SharePoint           → IT Device List, IT Device Changes
       → useDevices           → dashboard + table read the same fetch
```

Everything up to `ReviewGrid` is a pure function of its input. Nothing before the grid
touches React state, MSAL, or the network.

## 6. File layout

```
src/pages/DevicesPage.jsx              route: AppShell + stage switch
src/features/devices/
  parse/    labels.js          the 21 known labels + matching
            parseReport.js     text → { fields, unknownLabels, warnings }
            parseValues.js     ` | ` sub-structure parsers per field
            placeholders.js    junk/placeholder token table
  derive/   deriveDevice.js    parsed fields → typed device record
            deriveRam.js       sticks, installed vs reported, upgradability
            deriveStorage.js   drives, totals, SSD/HDD classification
            deriveCpu.js       vendor, model, generation, age band
            deriveIdentity.js  owner, department, device type
            deriveHealth.js    OS support, antivirus status
            riskScore.js       additive score + reasons + level
  sharepoint/
            deviceSchema.js    column definitions for both lists
            provisionLists.js  idempotent list + column creation
            syncDevices.js     read-all, diff, write, change rows
            readDevices.js     paged read + row → device record
  ui/       DropZone.jsx       drag & drop + file picker
            ReviewGrid.jsx     editable review table
            SaveProgress.jsx   progress + per-row results
            DeviceTable.jsx    the saved register, filtered by query string
            DeviceCharts.jsx   the CSS charts
            Leaderboards.jsx   the five ranked tables
  useDevices.js                the one SharePoint read for this section
src/styles/devices.css
```

Rationale for the split: `parse/` knows nothing about devices, `derive/` knows nothing
about SharePoint, `sharepoint/` knows nothing about React. Each layer is testable
without the one above it.

## 7. Parsing (`parse/`)

### 7.1 Normalisation

Strip BOM; `\r\n` → `\n`; strip non-breaking spaces and zero-width characters; trim
each line's trailing whitespace (the scan writes `AMD Ryzen 5 7430U with Radeon
Graphics` followed by nine spaces).

Structural lines, dropped before parsing: any line matching `^=+$`, and the exact
strings `COMPUTER INFORMATION` and `END OF REPORT`. Without this, the banner text
lands in `Remarks:` as an answer.

### 7.2 The 21 known labels

```
Name · Anydesk · Antivirus status · Remarks · Computer Name · Computer Model ·
Motherboard · Windows Version · Processor · GPU · Total RAM · RAM Slot Info ·
Storage Drives · Network Information · Antivirus · Monitor ·
PMW Server and credentials · Server folder · Microsoft Office · Adobe ·
Email data files found Active or Inactive account
```

Verified identical across all 17 sample files.

### 7.3 Block splitting — the rule that matters

**A line opens a new field only if the text before its first colon matches a known
label**, compared case-insensitively with internal whitespace collapsed.

This is load-bearing. A generic `^\w+:` split corrupts real data:

| Line | Naive split | Correct |
|---|---|---|
| `Total Slots: 2 \| Used Slots: 2` | new field "Total Slots" | value of `RAM Slot Info` |
| `Y: \| \\server\PMW\IT` | new field "Y" | value of `Server folder` |
| `C:\Users\User\...` inside an email path | new field "C" | part of the value |

### 7.4 Answer separation

1. Text after the colon on the label line, if non-empty, is the first answer. Covers
   `Antivirus status: NORTON ACTIVE` and the no-space variant
   `Antivirus status:NORTON NOT INSTALLED`.
2. Each following non-blank line is a further answer, until the next known label.
3. Blank lines are skipped, never treated as terminators. No sample has a blank line
   inside a block, and terminating on the next known label survives one if it appears.

A field's value is therefore always `string[]` — zero, one, or many answers. Nothing
downstream has to guess whether a field is single- or multi-valued.

### 7.5 Unknown labels

A line matching `^([A-Za-z][\w /&()+.'-]{0,60}):\s*(.*)$` that is not a known label is
recorded in `unknownLabels` — **unless** it contains ` | ` or begins with a single
drive letter and colon. Those two guards are what keep `Total Slots:` and `Y:` out.

Unknown labels are surfaced in review as "new field found in this file" and stored in
the `Extra Fields` JSON column. If the scan script gains a field later, it appears
rather than disappearing.

### 7.6 Sub-structure parsers (`parseValues.js`)

Each splits on ` | ` and trims. All return `null` for placeholder tokens (§7.7).

| Field | Shape | Notes |
|---|---|---|
| `RAM Slot Info` | `{ sizeGB, type, speedMhz, vendor, partNumber }[]` + `{ totalSlots, usedSlots }` | The summary line `Total Slots: N \| Used Slots: M` is recognised by prefix, not treated as a stick. `usedSlots` is empty in 5 of 17 files — fall back to `sticks.length`. |
| `Storage Drives` | `{ model, type, sizeGB }[]` | `type` of `Unspecified` means the scan could not read the media type; in every sample it is a mechanical disk. Mapped to `HDD (assumed)` and counted as mechanical, editable in review. |
| `Network Information` | `{ connection, ssid, ip, assignment }` | `Wi-Fi \| SSID: X \| IP: Y \| Dynamic`. The inner `SSID:` and `IP:` are stripped by prefix. |
| `Antivirus` | `{ product, enabled }[]`, de-duplicated | `AMIR-HP` lists HP Wolf Pro Security 22 times with conflicting states. De-duplicate by product name; a product is enabled if any of its entries is enabled. |
| `Motherboard` | `{ vendor, model }` | |
| `PMW Server and credentials` | `{ host, username }[]` | |
| `Server folder` | `{ letter, path }[]` | |
| `Adobe` | `{ product, version }[]` | |
| `Email data files found…` | `{ file, path, kind }[]` | `kind` is `mailbox` for `.ost`, `archive` for `.pst`. |
| `Microsoft Office` | `string[]` | Comma-separated on a single line, unlike every other multi-value field. |
| `GPU` | `string[]` | `VirtualMonitorDriver Device` is AnyDesk's virtual display, excluded from the real GPU list. |
| `Monitor` | `string[]` | `Default Monitor` is a Windows pseudo-device, excluded from the count. |

### 7.7 Placeholder tokens

Treated as null wherever they appear:

`None` · `Unknown` · `N/A` · `Manufacturer1` · `PartNum1` · `System Product Name` ·
`To Be Filled By O.E.M.` · `Default string` · `Not Specified` · empty string

`Manufacturer1` / `PartNum1` are unset SMBIOS strings on `DESKTOP-8SBR420`. Storing
them as a vendor name would produce a "Manufacturer1" category on the dashboard.

Note that `System Product Name` being present is itself a signal — it means the DMI
product string was never set, which in practice means a desktop assembled from parts.
It nulls the model but feeds device-type detection (§8.3).

## 8. Derivation (`derive/`)

### 8.1 RAM

- `installedRamGB` — sum of stick sizes. **This is the authoritative figure.**
- `reportedRamGB` — the number in `Total RAM`.
- `ramDiscrepancy` — true when they differ. `Total RAM` reports RAM *usable by
  Windows*, so an integrated GPU's reserved share is missing: `7 GB` is an 8 GB
  machine, `15 GB` is 16 GB. Ranking machines by the reported figure would place a
  16 GB laptop below an 8 GB one.
- `ramType` — most common across sticks (`DDR3` / `DDR4` / `DDR5`); `Unknown` when the
  scan could not read it.
- `ramSpeedMhz` — the minimum across sticks, since mixed sticks run at the slowest.
- `ramSlotsUsed` / `ramSlotsTotal` — from the summary line, `usedSlots` falling back to
  stick count.
- `ramUpgradable` — `ramSlotsUsed < ramSlotsTotal`. `EVONNE-HP` has one 8 GB stick in a
  two-slot board: fixable with one stick rather than a new machine.

### 8.2 Storage

`storageTotalGB` (sum), `driveCount`, `hasHdd` (any non-SSD), and `storageType`:
`SSD only` · `Mixed` · `HDD only` · `Unknown`.

### 8.3 CPU and device type

`cpuGeneration` from these patterns, in order:

| Pattern | Example | Generation |
|---|---|---|
| Core Ultra | `Intel(R) Core(TM) Ultra 5 125U` | `Ultra 1` |
| 5-digit Intel Core | `i7-1355U`, `i5-12400` | first two digits → 13, 12 |
| 4-digit Intel Core | `i7-3667U` | first digit → 3 |
| Pentium / Celeron / Atom, no generation | `Pentium(R) Dual CPU E2160` | none → obsolete |
| AMD Ryzen | `Ryzen 5 7430U` | series 7, **not** treated as an Intel-comparable generation |

AMD's mobile numbering does not map onto Intel's — a Ryzen 7430U is a Zen 3 part
wearing a 7000 badge. It gets a series number and its age band comes from RAM
generation and OS instead of a fabricated CPU generation.

`cpuAgeBand`: **Current** (Intel ≥ 10th gen or Core Ultra, or DDR5) · **Aging**
(Intel 7th–9th gen, or DDR4 with no generation reading) · **Obsolete** (Intel ≤ 6th
gen, DDR3 or older, or an ungenerationed Pentium/Celeron).

`deviceType` resolves in this order, because the computer name lies —
`DESKTOP-2A3ERS8` is an HP EliteBook laptop:

1. Motherboard model matches a known desktop board family (`PRIME`, `MS-7`, `P5G`,
   `PRO B`, `TUF`, `ROG`) → **Desktop**.
2. Computer model matches `Laptop|Notebook|Book|Pavilion|Inspiron|Latitude|Vostro|ThinkPad|IdeaPad|Precision \d{4}|Folio|Elite` → **Laptop**.
3. Model was the `System Product Name` placeholder → **Desktop** (unset DMI in
   practice means an assembled desktop).
4. Otherwise **Unknown**, flagged for review.

The verdict is always editable, and the review grid marks it as derived.

### 8.4 Identity

`owner` resolves through four sources in order, recording which one won in
`ownerSource`:

1. The `Name:` field, when filled (blank in all 17 samples, but it is the intended
   source).
2. A person's name in the filename bracket — `[QAQC FAIRUS]` → `Fairus`,
   `[QC SAM]` → `Sam`. Detected by removing tokens that match the known department
   list from the bracket contents; what remains, if anything, is a person.
3. The username in `PMW Server and credentials` — `server | ashraf` → `Ashraf`.
4. The local part of the first `.ost` mailbox — `lemon.cheong@pmw-group.com` →
   `Lemon Cheong` (dots and underscores become spaces, each word title-cased).

`department` is the bracket contents with the person's name removed, matched against
the departments seen so far: `ENGINEERING`, `FINANCE`, `SALES`, `QAQC`, `QC`,
`STOCKYARDF1`, `PML GUARDHOUSE`. Filenames without a bracket produce no department.

The bracket may be adjacent to the name with no space (`[QAQC FAIRUS]HPFL05_.txt`).
`computerName` falls back to the filename stem with the bracket and the trailing `_`
removed, when the `Computer Name` field is empty.

### 8.5 Health

- `windowsMajor` (`10` / `11`), `windowsEdition` (`Pro` / `Home` / `Home Single
  Language`), `osSupported` — false for Windows 10 and below, which reached end of
  support on 14 October 2025.
- `antivirusStatus` normalises the free-text `Antivirus status` field, which appears in
  seven spellings across the samples, and cross-checks it against the parsed
  `Antivirus` block:

  | Source text | Status |
  |---|---|
  | `NORTON NOT INSTALLED` | Not Installed |
  | `NORTON ACTIVATED`, `NORTON ACTIVE`, `NORTON INSTALLED (ACTIVE)` | Active |
  | `NORTON INSTALLED (DEACTIVATED)` | Installed — Inactive |
  | `NORTON INSTALLED (7 DAYS)` | Trial |
  | blank | derived from the `Antivirus` block, else Unknown |

- `avProtected` — true when any antivirus product is enabled. Windows Defender counts;
  a machine with Defender enabled and no Norton is protected, just not licensed.

### 8.6 Scan completeness

`scanComplete` is false when `Computer Name`, `Processor` and `Storage Drives` are all
empty — the signature of a scan that failed early, as in `CARMEN-HP_.txt`.

**Incomplete scans are excluded from every average, count and leaderboard on the
dashboard**, and appear only in the "re-scan needed" list. One failed scan otherwise
drags down the fleet's average RAM and inflates its risk.

### 8.7 Risk score (`riskScore.js`)

Additive. Each contributing signal is recorded in `riskReasons` so the dashboard can
explain a score rather than assert it.

| Signal | Points |
|---|---|
| `osSupported === false` | 40 |
| `antivirusStatus` is Not Installed, or `avProtected === false` | 30 |
| `installedRamGB <= 4` | 25 |
| `installedRamGB <= 8` (and > 4) | 15 |
| `cpuAgeBand === 'Obsolete'` | 25 |
| `cpuAgeBand === 'Aging'` | 10 |
| `hasHdd` | 10 |

`riskLevel`: **Critical** ≥ 60 · **High** 40–59 · **Watch** 20–39 · **OK** < 20.

Machines with `scanComplete === false` get a null score, not zero — an unscanned
machine is unknown, not healthy.

Sanity check against the samples: `DESKTOP-8SBR420` 100 (Win 10, Pentium E2160, 2 GB,
HDD), `HPFL05` 80 (Win 10, i7-3667U/DDR3, 8 GB), `AMIR-HP` 50 (Win 10, HDD),
`ASHRAF-PC` 15 (8 GB), `PMWL034` 0.

## 9. SharePoint

### 9.1 Provisioning

Two lists, created idempotently on first save, following the `ensureAssetList` /
`ensureAssetColumns` pattern in `src/services/sharePointService.js`.

**Do not copy `ensureColumns`** (the `IT Request Form` path): it never applies
`col.choices` to the request body, so its choice columns are created with no choices.
`ensureAssetColumns` does it correctly. That existing bug is out of scope to fix here.

Field creation rules:

| Kind | `FieldTypeKind` | `__metadata` type | Extra |
|---|---|---|---|
| Text | 2 | `SP.Field` | |
| Note | 3 | `SP.FieldMultiLineText` | `RichText: false`, `NumberOfLines: 6`, `AppendOnly: false` |
| DateTime | 4 | `SP.FieldDateTime` | **`DisplayFormat: 1`** (date *and* time) |
| Choice | 6 | `SP.Field` | `Choices: { results: [...] }` |
| Boolean | 8 | `SP.Field` | |
| Number | 9 | `SP.FieldNumber` | `DisplayFormat: 0` (integer) where applicable |

Two of these differ from what the existing service does, deliberately:

- **`DisplayFormat: 1` on DateTime.** The existing service hardcodes `0` (DateOnly) for
  every date column. Copying that would discard the time, which is the thing this
  feature was asked to get right.
- **`RichText: false` on Note.** A rich-text Note stores values wrapped in `<div>`
  markup, so multi-answer text would not round-trip.

The exact accepted body for `SP.FieldNumber` and `SP.FieldMultiLineText` creation is
to be confirmed against the live tenant during implementation; if a property is
rejected, fall back to `SP.Field` with the correct `FieldTypeKind` and record the
result in `AGENTS.md`.

### 9.2 `IT Device List` columns

`StaticName` uses no spaces (as `ASSET_REQUIRED_COLUMNS` does, avoiding the
`_x0020_` encoding of the older list). `Title` carries the readable name.

| StaticName | Title | Type |
|---|---|---|
| `Title` | Computer Name | Text (built-in) |
| `Owner` | Owner | Text |
| `OwnerSource` | Owner Source | Choice: Name field, Filename, Server credential, Email, Manual |
| `Department` | Department | Text |
| `DeviceType` | Device Type | Choice: Laptop, Desktop, Unknown |
| `ComputerModel` | Model | Text |
| `MotherboardVendor` | Motherboard Vendor | Text |
| `MotherboardModel` | Motherboard Model | Text |
| `AnydeskId` | AnyDesk ID | Text |
| `ScannedOn` | Scanned On | DateTime |
| `ImportedOn` | Imported On | DateTime |
| `ScannedOnMYT` | Scanned On (MYT) | Text |
| `SourceFileName` | Source File | Text |
| `WindowsVersion` | Windows Version | Text |
| `WindowsMajor` | Windows Major | Number |
| `WindowsEdition` | Windows Edition | Text |
| `OsSupported` | OS Supported | Boolean |
| `CpuModel` | CPU | Text |
| `CpuVendor` | CPU Vendor | Choice: Intel, AMD, Other |
| `CpuGeneration` | CPU Generation | Text |
| `CpuAgeBand` | CPU Age | Choice: Current, Aging, Obsolete, Unknown |
| `InstalledRamGB` | Installed RAM (GB) | Number |
| `ReportedRamGB` | Reported RAM (GB) | Number |
| `RamDiscrepancy` | RAM Discrepancy | Boolean |
| `RamType` | RAM Type | Text |
| `RamSpeedMhz` | RAM Speed (MHz) | Number |
| `RamSlotsUsed` | RAM Slots Used | Number |
| `RamSlotsTotal` | RAM Slots Total | Number |
| `RamUpgradable` | RAM Upgradable | Boolean |
| `RamSlotInfoRaw` | RAM Slot Info | Note |
| `StorageTotalGB` | Storage Total (GB) | Number |
| `DriveCount` | Drive Count | Number |
| `StorageType` | Storage Type | Choice: SSD only, Mixed, HDD only, Unknown |
| `HasHdd` | Has HDD | Boolean |
| `StorageDrivesRaw` | Storage Drives | Note |
| `AntivirusStatus` | Antivirus Status | Choice: Active, Installed — Inactive, Trial, Not Installed, Unknown |
| `AntivirusStatusRaw` | Antivirus Status (raw) | Text |
| `AntivirusProducts` | Antivirus Products | Note |
| `AvProtected` | Protected | Boolean |
| `NetworkType` | Network | Text |
| `Ssid` | SSID | Text |
| `IpAddress` | IP Address | Text |
| `IpAssignment` | IP Assignment | Choice: Dynamic, Static, Unknown |
| `GpuList` | GPU | Note |
| `MonitorCount` | Monitors | Number |
| `MonitorsRaw` | Monitors (raw) | Note |
| `MicrosoftOffice` | Microsoft Office | Note |
| `AdobeProducts` | Adobe | Note |
| `MappedDrives` | Mapped Drives | Number |
| `ServerFolders` | Server Folders | Note |
| `ServerCredentials` | Server Credentials | Note |
| `MailboxCount` | Mailboxes | Number |
| `ArchiveCount` | Archives | Number |
| `EmailDataFiles` | Email Data Files | Note |
| `RiskScore` | Risk Score | Number |
| `RiskLevel` | Risk Level | Choice: Critical, High, Watch, OK, Unknown |
| `RiskReasons` | Risk Reasons | Note |
| `ScanComplete` | Scan Complete | Boolean |
| `Remarks` | Remarks | Note |
| `ExtraFields` | Extra Fields | Note (JSON) |
| `RawReport` | Raw Report | Note |

`RawReport` holds the whole file — about 2 KB against the Note limit of 63,999
characters — so derivation rules can be re-run later without re-collecting files.

### 9.3 `IT Device Changes` columns

| StaticName | Title | Type |
|---|---|---|
| `Title` | Computer Name | Text |
| `FieldName` | Field | Text |
| `OldValue` | Old Value | Note |
| `NewValue` | New Value | Note |
| `ChangedOn` | Changed On | DateTime |
| `ChangedOnMYT` | Changed On (MYT) | Text |
| `ChangedBy` | Changed By | Text |
| `ChangeType` | Change Type | Choice: Added, Updated, Removed |

**Tracked fields only** — a change row is written for: `Owner`, `Department`,
`DeviceType`, `ComputerModel`, `WindowsVersion`, `OsSupported`, `CpuModel`,
`CpuAgeBand`, `InstalledRamGB`, `RamType`, `RamSlotsUsed`, `StorageTotalGB`,
`StorageType`, `AntivirusStatus`, `RiskLevel`.

`IpAddress`, `Ssid`, `MappedDrives` and the raw Note columns are **not** tracked. IP
addresses are DHCP-assigned and change constantly; logging them would bury real
hardware changes.

### 9.4 Malaysia time

MYT is a flat UTC+8. Three rules:

1. **Stored** as a true instant: `new Date(file.lastModified).toISOString()`. SharePoint
   keeps DateTime in UTC.
2. **Displayed in the portal** through `formatMYT`, which uses
   `Intl.DateTimeFormat('en-GB', { timeZone: 'Asia/Kuala_Lumpur' })` — correct
   regardless of the browser's or the site's timezone.
3. **Mirrored as text** into `ScannedOnMYT` / `ChangedOnMYT`, e.g.
   `19/08/2026 09:18 AM`. SharePoint renders DateTime columns in the *site's* regional
   timezone; if the site is not set to UTC+8 its own list view shows a different hour
   from the portal. The text mirror is unambiguous everywhere.

On first load the section reads `_api/web/regionalsettings/timezone` once. If the site
is not UTC+8, a dismissible notice states the difference and points at the mirror
column. The notice is informational — the portal's own figures are already correct.

**`malaysiaTime.js` needs two new styles.** It currently pins `hourCycle: 'h23'`, with
a comment warning that pairing `hourCycle` with `hour12` makes the engine discard the
former. AM/PM output therefore comes from new `datetime12` and `time12` styles that
pin `hourCycle: 'h12'` on the same principle — `h12` renders midnight as `12 AM`,
whereas `h11` would render it as `0 AM`. The existing 24-hour path is not modified.

### 9.5 Sync (`syncDevices.js`)

```
1. provisionLists()                       idempotent, once per save
2. readAllDevices()                       paged, $top=500, follow d.__next
3. index by Title (computer name), case-insensitive
4. for each reviewed record:
     absent  → insert
     present → diff tracked fields → update if changed → one change row per field
5. writes go through a 4-way concurrency pool
     429 or 503 → honour Retry-After, exponential backoff, 3 attempts
6. return per-row results { computerName, action, error }
```

Reading every row once and diffing in memory costs one request per 500 machines. The
alternative — a `$filter=Title eq '…'` per file — is one request per file and needs
`Title` indexed to survive the 5,000-item list view threshold.

Writes use a concurrency pool rather than SharePoint's multipart `$batch`. Hand-built
multipart bodies are error-prone, and 4-way concurrency with backoff imports 200
machines in well under a minute. `$batch` remains available as a later optimisation.

A save is **not** transactional. Partial success is normal and reported per row; the
retry re-sends only failed rows.

## 10. Screens

### 10.1 `/devices` — three stages on one route

`DevicesPage` renders `<AppShell title subtitle actions>` plus its own body, and does
not gate itself. Stage is component state, not a route.

**Drop.** A dropzone accepting multiple `.txt` files by drag or by picker — dragging
straight out of an Outlook attachment pane does not reliably produce files, so the
picker is not optional. Non-`.txt` files are rejected by name with a reason. Parsing
happens immediately, on the main thread: 200 files of 2 KB parse in well under the
time it takes to render the result, so no worker is needed here (unlike Data Studio,
where the workbooks are large).

**Review.** The table. One row per file, columns are the fields. A summary bar reads
`14 new · 3 updated · 1 needs attention`. Rows sort problems to the top: incomplete
scans, unknown labels, RAM discrepancies, unknown device type. Derived cells (Owner,
Department, Device Type, Scanned On) carry a marker and are editable inline; parsed
cells are read-only. Any row can be excluded from the save. A row whose computer name
matches an existing device shows the fields that will change.

**Save.** Progress bar with per-row results, failures listed with their reason and a
retry limited to failed rows.

### 10.2 `/devices/register` — the saved list

The register table, with filters in the query string exactly as `/requests` does:
`?risk=`, `?type=`, `?department=`, `?os=`, `?storage=`, `?ram=`. Sortable columns,
search by computer name or owner, CSV export of the current view.

### 10.3 `/devices/dashboard` — statistics

Six stat cards, each clicking through to the register with the matching filter:

**Total devices · Needs attention (Critical + High) · Unsupported OS · Unprotected ·
Average installed RAM · Stale scans (>180 days)**

Eight charts, in the CSS style of `DashboardPage.jsx`:

| Chart | Form | Answers |
|---|---|---|
| Risk mix | horizontal bars, danger palette | what needs attention |
| RAM distribution by size | horizontal bars | who has most and least |
| Laptop vs Desktop | horizontal bars | PC or laptop |
| OS mix | horizontal bars, Windows 10 in `--it-danger` | what is at risk |
| Storage type | horizontal bars | SSD, mixed, mechanical |
| CPU age band | horizontal bars | what is old |
| Devices per department | horizontal bars with average risk | where problems cluster |
| Scans per month | column chart, time axis | coverage over time |

Six leaderboards:

- **Highest RAM** — top 5
- **Lowest RAM** — bottom 5
- **Oldest hardware** — by CPU age band, then RAM generation
- **Recently scanned** — newest 5 by `ScannedOn`
- **Upgrade candidates** — `installedRamGB <= 8 && ramUpgradable`, the machines fixable
  with a stick rather than a replacement
- **Re-scan needed** — `scanComplete === false`, plus anything older than 180 days

Every figure links into the register with a query string, so a card and the rows it
opens read from the same `useDevices()` fetch and cannot disagree.

## 11. Error handling

| Failure | Behaviour |
|---|---|
| Non-`.txt` file dropped | Rejected by name, with the reason, before parsing |
| File unreadable | Named error for that file; the rest of the batch continues |
| No known label found in a file | "Not a device report" — excluded, not imported as an empty row |
| Core fields empty | Imported with `ScanComplete = false`, excluded from stats |
| Unknown label found | Row flagged, value stored in `ExtraFields`, import proceeds |
| Two files with the same computer name in one drop | Newest `Scanned On` wins; the other is flagged and excluded |
| Token acquisition fails at save | Handled by the session guard; the review table is preserved |
| List or column creation fails | Save aborts before writing any row, with the SharePoint message |
| Individual row write fails | Reported per row; retry re-sends only failures |
| 429 / 503 | Retry-After honoured, exponential backoff, 3 attempts |
| Site timezone is not UTC+8 | Dismissible notice; portal figures unaffected |

## 12. Testing

Vitest is already configured (`environment: 'node'`, `include: ['src/**/*.test.js']`)
with 73 tests passing. All new tests are pure-function tests in that environment; no
jsdom is added.

Written test-first:

- **`parseReport`** — the fixture set is the 17 real files, checked in under
  `src/features/devices/__fixtures__/`. Cases: inline value with and without a space
  after the colon; the `Total Slots:` and `Y:` label look-alikes; blank lines inside a
  block; banner and separator lines; the fully blank `CARMEN-HP` report; unknown label
  detection and its two guards.
- **`parseValues`** — each sub-parser: a stick line, the slots summary with an empty
  `Used Slots`, a drive with `Unspecified` type, the network line's inner `SSID:` and
  `IP:`, 22 duplicate antivirus entries collapsing to one, comma-separated Office,
  `VirtualMonitorDriver` and `Default Monitor` exclusion.
- **`deriveRam`** — installed vs reported divergence (7/8, 15/16), mixed stick speeds,
  `usedSlots` fallback, upgradability.
- **`deriveCpu`** — all five generation patterns, including `Core Ultra 5 125U`,
  `i7-3667U`, `i5-12400`, `Pentium E2160`, `Ryzen 5 7430U`.
- **`deriveIdentity`** — the four-owner-source chain in order, bracket with and without
  a space before the name, bracket containing both department and person, filename with
  no bracket, email local-part title-casing.
- **`riskScore`** — each signal in isolation, the four band boundaries at 20/40/60, and
  the five sample machines scoring 100/80/50/15/0. Incomplete scans score null.
- **`malaysiaTime`** — new `datetime12` / `time12` styles: midnight is `12:00 AM` not
  `0:00 AM`, noon is `12:00 PM`, and the existing 24-hour styles are unchanged.
- **`syncDevices`** — diffing against a fake existing-row map: insert, no-op, update
  with change rows, untracked field changing produces no change row, duplicate computer
  names in one batch.

Post-implementation: a browser pass through drop → review → save → dashboard against
the 17 real files, and a check that the created SharePoint columns have the intended
types.

## 13. Phasing

Each phase is independently verifiable, and nothing writes to SharePoint until phase 4.

1. **Parse** — `labels`, `parseReport`, `parseValues`, `placeholders`, with the 17
   files as fixtures. No UI.
2. **Derive** — RAM, storage, CPU, identity, health, risk. No UI.
3. **Drop and review UI** — dropzone, review grid, editing, exclusion. Still no
   SharePoint: the section is fully usable and demonstrable offline.
4. **SharePoint** — schema, provisioning, paged read, diff, write pool, change log.
5. **Register** — the saved table, filters in the query string, CSV export.
6. **Dashboard** — stat cards, charts, leaderboards, click-through filters.

Phases 1–3 deliver a working parser and review table that can be checked against real
files before a single SharePoint column is created.

## 14. Risks

| Risk | Mitigation |
|---|---|
| The scan script changes its label set | Unknown labels are captured and surfaced, never dropped; `RawReport` allows re-derivation |
| Device type guessed wrong | Motherboard-first ordering, `Unknown` rather than a coin flip, always editable |
| `File.lastModified` reset by copying | `Scanned On` is editable in review, and `Imported On` is always true |
| SharePoint field creation rejects a property | Documented fallback to `SP.Field` + `FieldTypeKind`; verified against the tenant in phase 4 |
| Site timezone not UTC+8 | Text mirror columns plus an in-app notice |
| Large drop is slow to write | 4-way pool with backoff; `$batch` available later if needed |
| 55 columns is unwieldy in SharePoint's own view | The portal is the primary reader; a default view can be trimmed in SharePoint without affecting the API |

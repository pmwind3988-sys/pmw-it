# Device map and machine lifecycle — design

**Date:** 2026-10-05
**Section:** `/devices`
**Status:** approved in conversation, awaiting spec review

## What this is for

IT wants to browse the fleet the way you would browse a game world, by entering a
department and seeing the machines in it. IT also wants to change who has a
machine without losing what came before. Three situations have to be recorded
truthfully:

1. **Someone gets a new laptop.** The new one becomes their machine. The old one
   keeps its own record and specs. It is not overwritten, and it is not deleted.
2. **An old laptop goes to somebody else.** That is the next line in its owner
   history, not a new record.
3. **A laptop is retired, and maybe brought back later.** It leaves the
   fleet figures but keeps its record. Bringing it back is one action.

Today the register knows a machine only by its computer name, and holds only
its current owner. It has no idea of a machine being spare or retired. A renamed
PC therefore looks like a new machine, and a reassignment simply overwrites the
previous owner.

## Decisions taken

| Question | Decision |
|---|---|
| What identifies a physical machine | Its **serial number**, read from the scan. The computer name is the fallback. |
| Where ownership lives | Inside `/devices`. It is not linked to the asset register's handovers. |
| A person shows up on a new machine | The import review **asks**: old one to stash, retire it, or keep both. |
| How browsing looks | An **auto-arranged world map**: locations around a central IT Stash hub, with a Graveyard at the edge. Enter a location to see its departments, enter a department to see its machines. |
| What the top level is | **Location** (F1, F3, PML, …), then department. |
| Where a machine's location comes from | The scan **file name** when it carries one, otherwise **set by hand**. A hand-set location beats the file. |
| Counting | Every level (location, department, IT Stash, Graveyard) shows its **laptop and desktop counts**. |

## 1. Data

### 1.1 `IT Device List` — four new columns

| Column | Kind | Meaning |
|---|---|---|
| `Location` | text | Site code: `F1`, `F3`, `PML`, … Blank means "no location yet". |
| `SerialNumber` | text | Manufacturer serial from the scan. Blank when the report has none, or when the value is a placeholder. |
| `Status` | choice | `In use`, `In repair`, `Spare`, `Retired`. |
| `StatusChangedOn` | datetime (`DisplayFormat: 1`) | When the status last changed. |

- **No migration:** a blank `Status` reads as `In use`, so every existing row is correct the moment the column exists.
- **`Owner` and `Department` mean "who has it NOW".** A `Spare` or `Retired` machine has both blank. Its last holder lives in the owner history.
- **Adding to `Status` later:** the column is reconciled through `mergeChoices`, which only ever adds options, the same as every other choice column.
- **`Location` is text, not a choice column, on purpose.** A new site must not need a column change before a machine can be put there.
- **Which locations exist** is worked out the same way `assets/categories.js` works out categories: a short built-in list (`F1`, `F3`, `PML`) plus every location a row is actually using. There is no second list to disagree with the rows. Adding one is typing a new code in the location field.
- **A location shows by its code.** Full names (such as "Factory 1") can be added to the built-in list later without touching any row.
- **`location` joins the hand-editable fields.** It sits beside owner, department and device type in `EDITABLE_FIELDS`, so it gets the same `ManualFields` protection against re-import.

### 1.2 `IT Device Assignments` — new list, one row per stint

| Column | Kind | Meaning |
|---|---|---|
| `Title` | built in | Computer name at the time (readable in SharePoint without a join). |
| `DeviceId` | number | The `IT Device List` row id. This, not the name, ties the stint to the machine. |
| `Owner` | text | Who had it. |
| `Location` | text | Their location during this stint. |
| `Department` | text | Their department during this stint. |
| `AssignedOn` | datetime | Start. |
| `AssignedOnApprox` | boolean | True when the start is only known as "since at least" (see 1.4). |
| `EndedOn` | datetime | End. Blank means this is the current stint. |
| `EndReason` | choice | `Reassigned`, `Replaced`, `To stash`, `To repair`, `Retired`. |
| `Note` | note (`RichText: false`) | Free text IT typed in, e.g. "screen cracked". |
| `RecordedBy` | text | Signed-in user who made the change. |

A machine has **at most one open stint**, and only while its status is `In use`
or `In repair`. A machine in repair keeps its stint open: it is still that
person's laptop, just away being fixed. Any repair time worth showing comes from
the status history in the change log, not from a separate stint.

### 1.3 `IT Device Changes` — one new column

- **`DeviceId` (number)** is added, so the spec history follows the machine through a rename.
- **Older rows have no `DeviceId`.** They are matched by `Title`, the old behaviour, as a fallback.
- **New change types are logged:** `computerName` (a rename), `serialNumber` (first seen), `location` and `status`. `status` also lands in the change log, so "when was it in repair" can be answered.

### 1.4 Machines that predate this feature

**They have no stints.** The first lifecycle action on such a machine writes the
outgoing owner's stint before closing it:

- **Start date:** `AssignedOn` = the row's SharePoint `Created`, with `AssignedOnApprox: true`.
- **On screen** it reads "since at least 21 Aug 2026".

Nothing is written for machines nobody touches.

## 2. Import

### 2.1 Reading the serial

**New scan label.** The scan script gains one line, and `parse/labels.js` gains
the label `Serial Number`:

```powershell
"Serial Number: $((Get-CimInstance Win32_BIOS).SerialNumber)"
```

**Placeholder serials.** Home-built and some OEM desktops report junk instead of a
real serial. These count as no serial:

- the existing `isPlaceholder` set
- `System Serial Number`
- `Chassis Serial Number`
- `0`
- `Default string`
- any value made of one repeated character

The list lives beside `isPlaceholder`, as `isPlaceholderSerial` in
`parse/placeholders.js`.

**Normalising.** Serials are compared upper-cased with whitespace removed, the same
rule as `normaliseCode` in assets.

### 2.1a Reading the location from the file name

**The convention.** IT already puts the department in a bracket:
`[ENGINEERING] AMIR-HP.txt`. The location goes first inside the same bracket:

- `[F1 ENGINEERING] AMIR-HP.txt` → location `F1`, department `Engineering`
- `[PML FINANCE] EVONNE-HP.txt` → location `PML`, department `Finance`

**How it is matched (`derive/deriveLocation.js`, pure).**

- The first word of the bracket is checked against the known locations (see 1.1), ignoring case. A match is the location, and the rest is the department as it is parsed today.
- No match: there is no location, and the whole bracket is the department, exactly as today. So `[ENGINEERING] X.txt` keeps working.

**The two names the importer already knows** are split the same way:

| Today's department | Location | Department |
|---|---|---|
| `STOCKYARDF1` | `F1` | `Stockyard` |
| `PML GUARDHOUSE` | `PML` | `Guardhouse` |

A location code glued onto the end of a known department name (`STOCKYARDF1`) is split off only when the code is on the known list. That stops a department that happens to end in a code-like suffix from being cut.

**On re-import:**

- A file with no location never clears one that is set.
- A hand-set location (in `ManualFields`) is never overwritten.
- A file location that differs from a non-manual one updates it, logged as a change.

**Rows already in SharePoint** keep their old department text until the reports are imported again or the location is set by hand. Their location stays blank until then, and the map shows them under **No location yet**.

### 2.2 Matching an incoming report — `lifecycle/matchIncoming.js` (pure)

The rules run in this order:

1. **Usable serial, and a register row with that serial:** it is that machine.
   - If the computer name differs, the machine was **renamed**. Update `Title` and log a `computerName` change.
   - The match holds whatever the row's status is. A retired machine that is scanned again is a machine being **brought back**, and the review says so (see 2.4).
2. **Otherwise, a register row with the same computer name:**
   - **That row has a different usable serial:** this is a **different machine reusing the name**. Treat it as new, and flag it in review: *"PMWL034 was a different machine (serial 5CG1…)."*
   - **That row has no serial:** it is the same machine, and the serial is written on it now.
3. **Otherwise:** a new machine.

Matching is planned for the whole batch at once, so two reports in one drop that
claim the same serial are flagged rather than both written.

### 2.3 Replacement prompt — `lifecycle/replacements.js` (pure)

**When it appears.** For every incoming machine whose owner already holds a
**different** `In use` / `In repair` machine in the register, the review row gets
a required choice:

> **Carmen already has CARMEN-HP (Laptop, in use).** Is this a replacement?
> ( ) Old one to IT Stash   ( ) Retire old one   ( ) Keep both

**How owners are compared.** Owner names are compared trimmed and case-folded.

**If Carmen has two old machines.** One prompt is shown per old machine.

**Save waits for an answer.** The Save button stays disabled and says
"2 replacements need an answer" until every prompt is answered. The rows needing
an answer sort to the top of the review, the same as flagged rows today.

**What each answer writes.** On save, the old machine's open stint closes with
`EndReason: Replaced`, and the old machine becomes `Spare` or `Retired` with its
owner and department cleared.

### 2.4 Owner changes the scan reveals

**A scan names a new owner on a matched machine.** Unless `owner` is in
`ManualFields`, this is a reassignment. The open stint closes as `Reassigned`
and a new one opens, starting at the scan date.

**A hand-typed owner still wins.** `applyManualOverrides` is unchanged, and the
stint is untouched.

**A machine scanned while `Spare` or `Retired`.** It returns to `In use` with
the scanned owner, and the review shows a notice: *"Brought back from IT Stash"*.

### 2.5 Order of writes

**The order:**

1. Device rows (insert/update).
2. Stints.
3. Change log.

**What a failure leaves behind.**

- **Stints fail:** a machine can carry a status its history does not explain. That is recoverable by retrying.
- **The device write fails:** no stint is written for that machine.

**What the result reports.** The existing result summary gains the counts of
stints written and stints failed.

## 3. Screens

### 3.1 Map tab — `/devices?view=map` (new default)

**Tabs.** The tab bar becomes **Map · Dashboard · Register · Import**. A bare
`/devices` opens the Map.

**Three levels, one address each.** Back and shared links work at every level.

| Level | Address | Shows |
|---|---|---|
| World | `?view=map` | Locations, the IT Stash, the Graveyard |
| Location | `?view=map&location=F1` | That location's departments |
| Department | `?view=map&location=F1&department=Engineering` | That department's machines |

**Tiles on the world map.**

| Tile | Holds |
|---|---|
| Each location | Its `In use` and `In repair` machines |
| **No location yet** (dashed) | In-use machines with a blank location. Shown only when non-empty. |
| **IT Stash** (centre hub) | `Spare` machines, wherever they last were |
| **Graveyard** (bottom edge) | `Retired` machines |

**Tiles inside a location.**

- One tile per department.
- A centre tile for the location itself, with its totals.
- An **Unassigned** tile for its machines with no department, shown only when non-empty.

**What every tile shows, at every level:**

- its name (a location shows its code large, e.g. **F1**)
- the machine count
- **laptops and desktops counted separately**, each with its glyph. `Unknown` device types are counted as "other" only when there are any.
- a health bar: Critical / Needs attention / Moderate / Optimal, as proportions
- a badge for the critical count, or "All clear"
- on a location tile only: its biggest departments as chips (*Engineering 14 · Logistics 12 · +4 more*)

The IT Stash and Graveyard tiles carry the same laptop and desktop split.

**Layout (`map/mapLayout.js`, pure and tested).** The same function lays out both
levels. Given the tiles, it returns grid positions:

- The centre tile (the IT Stash, or the location's own tile) sits in the middle.
- The other tiles ring it, largest headcount first.
- On the world map, the Graveyard spans the bottom.

Paths between tiles are drawn as decoration only. A ring that would exceed the
grid adds a row, so any number of locations or departments fits.

**Responsive.** Below 768px each level renders as a vertical list of the same
tiles. On the world map the IT Stash comes first, then the locations, then the
Graveyard, then No location yet.

**Keyboard and screen readers:**

- Arrow keys move focus between tiles, by grid position.
- Enter or a click goes one level in.
- Esc goes one level out.
- Focus is visible.
- Each tile is a link labelled in full, e.g. *"F1, 65 machines, 41 laptops, 24 desktops, 8 critical"*.

**Motion.** Going in plays a short zoom. With `prefers-reduced-motion`, it plays
no animation.

**Scope.** The dashboard's department scope picker does not apply here. The map
always shows the whole fleet.

### 3.2 Inside a department — `?view=map&location=F1&department=Engineering`

**Header:**

- a breadcrumb, *Map › F1 › Engineering*
- a level-title header: department name, the persona it is judged against, the health bar, and counts (machines, laptops, desktops, critical, needs attention, in repair)
- Esc, or the breadcrumb, goes back to the location

**Body:**

- **Search box.** It searches owner, computer name and serial within the department.
- **Machine cards, grouped under *In use* and *In repair*.** The IT Stash, the Graveyard and No location yet open straight to a card list with a single group, since they have no departments inside.
- **Each card shows:**
  - owner, or "In stash" / "Retired 3 Mar 2026"
  - computer name
  - laptop or desktop glyph
  - three short bars, for CPU generation, RAM and storage, filled against fixed full scales (15th-gen CPU, 32 GB, 1 TB) so cards compare at a glance, and coloured by the fit verdict
- **A card opens the machine page.**

**A department name at two locations is two separate tiles.** Finance at F1 and
Finance at PML are separate tiles, each with its own counts.

### 3.3 Machine page — `/devices/:id` (existing, extended)

**Added to the top:**

- a status badge
- the current owner, location and department
- an action bar:

| Action | Shown when | Writes |
|---|---|---|
| **Change owner** | In use, In repair | Closes the stint as `Reassigned`. Opens a new one for the new owner, location and department. |
| **To repair** | In use | Status → `In repair`. The stint stays open. |
| **Back in use** | In repair | Status → `In use`. |
| **To stash** | In use, In repair | Closes the stint as `To stash`. Status → `Spare`. Clears the owner. |
| **Retire** | any but Retired | Closes any stint as `Retired`. Status → `Retired`. Asks first, through `useConfirm`. |
| **Assign / Bring back** | Spare, Retired | Status → `In use`. Opens a stint for the chosen owner. |

**The assignment dialog** has:

- **owner:** free text, suggesting owners already in the register
- **location:** picks from the known locations, or takes a new code
- **department:** suggests departments already used at that location
- **effective date:** defaults to today, and is parsed at local noon as `parseFormDate` does
- **an optional note**

**New sections:**

- **Owner history.** A timeline of stints, newest first, with dates, end reason and note. "since at least" marks an approximate start.
- **Spec history.** Change-log rows for this machine, grouped by date, e.g. *RAM 8 → 16 GB*. Renames show inline as *Renamed CARMEN-HP → AISYAH-HP*.

**The existing spec groups** stay below, unchanged.

### 3.4 Register and dashboard

**Register:**

- It gains **Location** and **Status** columns, with a filter for each.
- **Retired is hidden by default.** The filter's header says so while it is.
- **Owner edits write history.** Editing `owner` or `department` inline there goes through the same lifecycle write as Change owner, so the register and the history cannot disagree.

**Fleet figures** count `In use` + `In repair` only:

- dashboard stat cards
- charts
- heatmap
- leaderboards

**IT Stash count.** The dashboard gains a stat card for spare machines, which links to the IT Stash.

## 4. Where the code goes

```
src/features/devices/
  derive/
    deriveLocation.js    location out of the file-name bracket (pure)
  lifecycle/
    status.js            STATUSES, statusOf(row) (blank → In use), counts-in-fleet
    matchIncoming.js     serial-then-name matching, renames, name reuse
    replacements.js      which incoming machines displace whom
    planLifecycle.js     an action → { deviceBody, closeStint?, openStint?, changes }
    stints.js            open stint of a machine, legacy "since at least" stint
  sharepoint/
    assignmentSchema.js  ASSIGNMENT_LIST_NAME, columns, to/from list item
    readAssignments.js   all stints, paged like readDevices
    writeLifecycle.js    performs a planLifecycle result, writes in order
  map/
    mapLayout.js         zones → grid positions (pure)
    zones.js             devices → tile summaries for either level, incl. laptop/desktop counts (pure)
    locations.js         built-in locations + those in use (pure)
  ui/
    WorldMap.jsx         the map tab
    LocationView.jsx     inside a location
    DepartmentView.jsx   inside a department
    MachineCard.jsx
    LifecycleActions.jsx action bar + dialog
    OwnerHistory.jsx
    SpecHistory.jsx
    ReplacementPrompt.jsx  the choice in a review row
```

- `deviceSchema.js` gains the four columns. `updateDevice.js` adds `location` to `EDITABLE_FIELDS`. `deriveIdentity.js` hands its bracket to `deriveLocation` before reading the department.
- `CHANGE_COLUMNS` gains `DeviceId`.
- `provisionLists.js` declares the third list.
- `syncDevices.js` uses `matchIncoming` in place of `indexByName`, and performs the replacement answers and scan-revealed reassignments.
- Styles go in `styles/devices.css`.
- Glyphs (map, stash, graveyard) are added to `ui/Icons.jsx`, which must be imported wherever `NAV_ITEMS` or a component references them.

**Layering holds:**

- `lifecycle/` and `map/` import no React and no SharePoint.
- `sharepoint/` imports no React.

## 5. Errors

- **A lifecycle action that fails** leaves the page as it was and shows the error with Retry. Nothing is half-applied on screen.
- **Write order.** The device row is written first, then the stint. If the stint fails after the row landed, the page says so and offers to write the missing stint. That is the same operation re-run, which is safe because closing an already-closed stint is a no-op.
- **Two people changing the same machine at once.** The action re-reads the machine and its open stint immediately before planning, the same rule as `commitHandover`. If the status changed in between, it refuses with *"This machine was just moved to IT Stash by someone else"* instead of writing over it.

## 6. Testing

**Pure modules are written test-first:**

| Module | What the tests pin |
|---|---|
| `matchIncoming` | serial hit, rename, name reuse with a different serial, serial-less fallback, placeholder serials, duplicate serial in one batch, retired machine rescanned |
| `replacements` | one old machine, two old machines, owner case and space differences, keep-both, a machine already spare is not prompted |
| `planLifecycle` | every action in the table, legacy stint synthesis, close-is-idempotent, refusal on stale status |
| `status` / fleet filter | blank → In use, Retired and Spare excluded from figures |
| `mapLayout` | 0, 1, 5 and 20 tiles, centre tile centred, no overlaps, deterministic order, same function for both levels |
| `zones` | No location yet and Unassigned only when non-empty, badge counts, laptop/desktop/other split, same department at two locations kept apart |
| `deriveLocation` | `[F1 ENGINEERING]`, `[ENGINEERING]` unchanged, `STOCKYARDF1`, `PML GUARDHOUSE`, unknown first word, case |
| `locations` | built-in plus in-use, no duplicates by case, blank ignored |
| `zones` | Unassigned only when non-empty, badge counts |
| `assignmentSchema` | round-trip, `DateOnly` not used, note not rich text |

**The existing suite** must stay green, notably `syncDevices`, `syncPhases`,
`provisionLists` and `updateDevice`.

**Checks run:** `npm run build`, `npm run lint` (no new errors beyond the known
`ThemeContext` one), and the full test suite.

**In the browser:** the screens need a Microsoft sign-in to load real rows. They
are checked in the dev server against the sample reports where possible, and the
real-data check is left to the user.

## Out of scope

- Linking a machine to an asset-register item or handover.
- Signatures on a reassignment.
- A hand-drawn site map. The layout is generated.
- Bulk lifecycle actions from the register, such as retiring several at once.
  The register's existing multi-select is for removal only.

# Device standards — admin-set grades per part

**Date:** 2026-10-06
**Section:** `/devices` (builds on the device map and lifecycle branch, `devices-map-lifecycle`)
**Status:** approved in conversation, awaiting spec review

## What this is for

IT wants to decide, in the portal and without a developer, what counts as Critical, Needs
attention, Moderate and Optimal. Today those thresholds are numbers written into
`derive/persona.js` and rules written into `derive/deviceFit.js`, and they produce one verdict
for a whole machine.

After this change:

- **Each part of a machine gets its own grade**, judged against the standard for the machine's
  workload profile.
- **There is no overall machine grade any more.**
- **Admins set the standards** and the four grade colours on a settings page.

## Decisions taken

| Question | Decision |
|---|---|
| Scope of a standard | Per **workload profile**: Engineering / Technical / Media, Desk, Field. Each **part** has its own scale. |
| Overall machine grade | **None.** Only per-part grades. |
| Parts graded | **CPU, RAM, Storage, Graphics, Windows.** |
| Not graded any more | Office licence, antivirus and network link. They stay on the machine page and in the separate risk score, which is unchanged. |
| Map tiles | **One thin bar per part**, split by grade, plus a badge "N machines with a critical part". |
| Who may edit | **SharePoint decides.** Whoever has edit rights on the standards list can save; everyone else sees the page read-only. |
| Colours | **The admin picks one colour per grade** (Critical, Needs attention, Moderate, Optimal, Unknown). Colour in the device section only ever means a part's grade, never a machine or a person. |

## 1. The standard

### 1.1 Grades

The four grades are `Critical`, `Needs attention`, `Moderate` and `Optimal`.

A fifth state, `Unknown`, means the scan did not report the value. It is never a guess.

### 1.2 Cut-offs: CPU, RAM and storage size

**The rule.** A numeric part is graded by three numbers, `criticalBelow`, `attentionBelow` and
`optimalFrom`:

| Value | Grade |
|---|---|
| `< criticalBelow` | Critical |
| `< attentionBelow` | Needs attention |
| `< optimalFrom` | Moderate |
| `≥ optimalFrom` | Optimal |

**What a save must satisfy.** `criticalBelow ≤ attentionBelow ≤ optimalFrom`.
- If two numbers are equal, the band between them is empty. That is allowed. For example,
  Desk RAM has no "Needs attention" band.
- Numbers out of order are refused on save, with the reason shown next to the field.

**How each part is graded:**

- **CPU** is graded on the Intel-equivalent generation the importer already computes
  (`cpuGenerationRank`), so AMD lands on the same scale.
- **CPU families with no generation** count as Critical whatever the cut-offs. These are
  Pentium, Celeron and AMD before Zen (`cpuAgeBand === 'Obsolete'`).
- **A CPU with no rank** that is not obsolete is `Unknown`.
- **RAM** is graded on installed GB (`installedRamGB`, the summed slots, not Windows' usable
  figure).
- **Storage size** is graded on `storageTotalGB`, which already leaves out the scan's own USB disk.

### 1.3 Choice grades: storage type, graphics and Windows

**Storage type.** One grade for each of `HDD only`, `Mixed` and `SSD only`.
- **The storage grade is the WORSE** of the type grade and the size grade.
- A 128 GB SSD can therefore be Needs attention, and a 1 TB hard disk is still Critical.

**Graphics.** One grade for `Dedicated card` and one for `Built-in graphics`, using the existing
`dedicatedGpu` flag. A machine with no graphics information is `Unknown`.

**Windows.** One grade for each of:
- `Windows 11`
- `Windows 10 (still supported)`
- `Out of support` (any version with `osSupported === false`)
- `Other / older`

A missing version is `Unknown`.

### 1.4 Which profile a department uses

**The mapping.** The standard holds a department → profile map. It starts from today's
`BY_DEPARTMENT` table in `persona.js`.

**Unmapped departments.** A department not in the map, and a machine with no department, use
the **Desk** profile, as they do today.

**Editing the map.** The settings page lists every department the register actually holds, each
with a profile picker.

### 1.5 Colours

**What is stored.** Five colours, one per grade plus Unknown, saved with the standard.

**Text on coloured chips.** It is black or white, chosen automatically by the colour's brightness
so it stays readable.

**Unreadable choices.** If the admin picks two grade colours too alike to tell apart (contrast
ratio under 1.5 between any two), the page warns but still allows the save.

**Dark mode.** The same colours are used, with the text rule above.

**Where they apply.** Only inside the device section. Nothing else in the portal changes colour.

### 1.6 Default standard (the first version, used until somebody saves one)

The defaults reproduce today's behaviour where it existed, and propose numbers where it did not.
The proposed numbers are marked †.

| Part | Engineering / Technical / Media | Desk | Field |
|---|---|---|---|
| CPU generation: critical below / attention below / optimal from | 8 / 10 / 12 † | 6 / 8 / 10 † | 7 / 10 / 12 † |
| RAM GB: critical below / attention below / optimal from | 8 / 16 / 32 | 8 / 8 / 16 | 8 / 16 / 16 |
| Storage GB: critical below / attention below / optimal from | 128 / 256 / 512 † | 128 / 256 / 256 † | 128 / 256 / 512 † |
| Storage type: HDD only / Mixed / SSD only | Critical / Needs attention / Optimal | same | same |
| Graphics: dedicated / built-in | Optimal / Needs attention | Optimal / Optimal | Optimal / Optimal |
| Windows: 11 / 10 supported / out of support / other | Optimal / Moderate † / Critical / Critical † | same | same |

| Grade | Default colour |
|---|---|
| Critical | `#dc2626` |
| Needs attention | `#f59e0b` |
| Moderate | `#1a88de` |
| Optimal | `#12a150` |
| Unknown | `#8a97a8` |

## 2. Storage and permissions

### 2.1 `IT Device Standards` list (new, append-only)

| Column | Kind | Meaning |
|---|---|---|
| `Title` | built in | `v<n>`, the version number |
| `Standard` | note (`RichText: false`) | The whole standard as JSON: profiles, the department map and the colours. Has a `schema: 1` field. |
| `SavedOn` | datetime (`DisplayFormat: 1`) | When it was saved. |
| `SavedBy` | text | The signed-in user. |
| `Summary` | note | What changed, in words. For example: "Engineering RAM optimal from 32 → 24 GB; Critical colour changed". |

**Every save adds a row.** Nothing is overwritten.

**The current standard** is the row with the highest `Id`.

**History.** The settings page lists the last 20 versions, each with who saved it, when, and the
summary. Restoring an old version saves it again as a new row, so the history never loses a step.

**Reading the list.** It is read with one request: `$orderby=Id desc&$top=20`.

### 2.2 Who may edit

**The check.** On opening the page, the portal asks SharePoint whether the user may add items to
`IT Device Standards`, through `EffectiveBasePermissions` on the list (the `AddListItems` bit).

**Without the right**, the page is read-only. A note says "Only people with edit rights on the IT
Device Standards list can change these. Ask IT."

**SharePoint is the real lock.** A save attempt without rights is refused by SharePoint whatever
the page shows, and the page reports it.

**The list does not exist yet.** It is created on the first save, by the same `provisionSchema`
rules as every other list. Until then:
- everyone reads the default standard;
- the page treats the user as able to edit if they can create lists on the site;
- otherwise it is read-only.

### 2.3 A bad or unreadable standard

**The rule.** If the latest row's JSON does not parse, or fails validation, the portal falls back
to the newest valid version, else the defaults.

**What the page shows.** A visible note naming the version it skipped. Grading never stops
because of a bad row.

## 3. Grading

**`derive/partGrades.js`** (pure, new) takes `(device, standard)` and returns:

```
{
  profileKey, profileLabel,
  parts: {
    cpu:      { grade, value: '8th gen', reason: 'Below the 10th gen this work needs for Moderate' },
    ram:      { grade, value: '16 GB',   reason: … },
    storage:  { grade, value: '256 GB SSD only', reason: … },
    graphics: { grade, value: 'Built-in', reason: … },
    windows:  { grade, value: 'Windows 11 Pro', reason: … },
  },
  criticalParts: ['ram', …],
  attentionParts: [ … ],
  actionRequired: 'Upgrade now: RAM, Storage' | 'Plan: CPU' | 'Nothing to do' | 'Re-run the scan',
}
```

**Each part's `reason`** is one sentence naming the threshold it was judged against.

**`actionRequired`:**
- If any part is Critical, it names the Critical parts.
- Else, if any part is Needs attention, it names those.
- Else it is "Nothing to do".
- An incomplete scan gives "Re-run the scan" and every part `Unknown`.

**Computed on read, never stored.** Grades are computed when the device list is read, with the
current standard. A saved standard therefore regrades the whole fleet on the next page load,
with no re-scan and no column change.

**What `deviceFit.js` becomes.**
- Its fleet-wide verdict and its fields (`fitStatus`, `fitReasons`) are removed.
- The form-factor suggestion moves into `partGrades` unchanged.
- `riskScore.js` is not touched.
- `persona.js` keeps the profile labels and blurbs. Its numbers move into the default standard
  (`standards/defaultStandard.js`).

## 4. Screens

### 4.1 Standards tab — `/devices?view=standards`

**Layout:**
- One column per profile.
- Sections for CPU, RAM, Storage, Graphics and Windows, each with its fields as described in 1.2
  and 1.3.
- Then **Departments**: every department in the register, each with a profile picker.
- Then **Grade colours**: five colour pickers with hex fields, a live sample chip for each, and
  "Reset to defaults".

**Live preview.** While editing, a panel shows how many machines would change grade, per part.
For example, "RAM: 12 machines → Needs attention, 3 → Optimal". It is computed from the loaded
register with the unsaved standard. Nothing is saved until **Save standard**.

**Save** asks for confirmation through `useConfirm`. The dialog shows the change summary and the
preview counts. The save button uses `Button loading`.

**History** sits below: the last 20 versions, each with a **Restore** button. Restore asks first.

**Read-only users** see the same page with every field disabled and the note from 2.2.

**Accessibility.** Inputs are at least 16px and every touch target at least 44px. Each colour
picker has a text label and a hex input.

### 4.2 Everywhere else in the device section

**Map tiles** (world and location). The single health bar is replaced by **five thin bars**:
CPU, RAM, Storage, Graphics and Windows.
- Each bar is split by how many of the tile's machines have that grade on that part.
- Bars use the grade colours, and each has its part label.
- The badge reads "N machines with a critical part", or "All clear".
- The tile's screen-reader label names the critical counts per part.

**Department list.** Machines are sorted by number of Critical parts, then Needs attention parts,
then name.

**Machine cards:**
- Five grade chips (CPU, RAM, Storage, Graphics, Windows), each in its grade colour.
- The coloured top border and the single verdict pill go away.
- The CPU, RAM and storage bars stay, coloured by that part's own grade.

**Machine page.** The "Fit for the work" card becomes **"Parts against the [profile] standard"**:
- a five-row table of part, value, grade chip and reason;
- the action sentence;
- the form-factor note.

**Dashboard:**
- "Not fit for the work" becomes **"Machines with a critical part"**. It opens the register with
  a filter for that.
- The department heatmap becomes department × part, each cell the number of machines Critical
  on that part.
- Other cards are unchanged.

**Register:**
- The "Device health" and "Why" columns are replaced by five part columns, each showing the
  grade chip.
- The filter `fit=` is replaced by `part=<part>:<grade>` (for example `part=ram:Critical`) and
  `critical=1` (any critical part).
- An old link with `fit=` is ignored rather than erroring.

**Colour in the device section** comes from the grade colours, through CSS custom properties
(`--grade-critical`, `--grade-attention`, `--grade-moderate`, `--grade-optimal`,
`--grade-unknown`, plus a `-ink` text colour for each). They are set on the device section's
root element. The device section no longer uses hard-coded grade colours anywhere.

## 5. Where the code goes

```
src/features/devices/
  standards/
    defaultStandard.js   the first version (section 1.6) and the grade list
    validateStandard.js  ordering, required fields, schema version (pure)
    diffStandard.js      the change summary sentence (pure)
    previewChanges.js    machines that would change grade, per part (pure)
    gradeColors.js       ink colour from luminance, contrast warning, CSS vars (pure)
  derive/
    partGrades.js        device + standard → per-part grades (pure); replaces deviceFit.js
  sharepoint/
    standardsSchema.js   list name, columns, to/from item
    readStandards.js     latest valid + history (last 20)
    saveStandard.js      append a version; permission check
  useStandards.js        loads the standard once per page, exposes it and canEdit
  ui/
    StandardsPage.jsx    the tab
    PartBars.jsx         the five map-tile bars
    PartChips.jsx        the five card chips
    PartTable.jsx        the machine page table
```

**How the standard reaches the screens.** `useDevices` grades rows with the standard from
`useStandards`, which every device screen already reads through.

**Layering holds:**
- `standards/` and `derive/` are pure.
- `sharepoint/` imports no React.

## 6. Errors

**SharePoint unreachable while reading the standard.** The section grades with the defaults and
shows a quiet banner: "Using the default standard; the saved one could not be read".

**A refused save** (permissions, or a network failure) keeps the edits on screen and offers Retry.
Unsaved edits are never discarded.

**Two admins saving at once.** Both succeed as consecutive versions. The second save's summary
is computed against what that admin loaded, and the history shows both.

## 7. Testing

| Module | What the tests pin |
|---|---|
| `partGrades` | every band boundary (`<` vs `≥`), empty bands, obsolete CPU, unknown values, storage = worse of type and size, every Windows bucket, unmapped department → Desk, `actionRequired` text, incomplete scan |
| `validateStandard` | order refused, missing part refused, unknown schema refused |
| `diffStandard` | readable summary of numeric, choice, mapping and colour changes |
| `previewChanges` | counts per part and grade |
| `gradeColors` | ink choice on light and dark colours, too-similar warning |
| `readStandards` | newest valid wins; a bad newest falls back with a note; a missing list gives the defaults |
| `saveStandard` | appends, never updates; refusal surfaced |

**Existing tests that change:** `deviceFit` and `deviceStats` (fit counts), zones (the per-part
bars), and the register filters. These are rewritten against the new shape, not deleted.

## Out of scope

- Adding, renaming or removing workload profiles. There are three.
- Grading Office licence, antivirus or network.
- An overall machine grade of any kind.
- Grade colours outside the device section.
- Per-person overrides, for example "this machine is fine for this person despite the standard".

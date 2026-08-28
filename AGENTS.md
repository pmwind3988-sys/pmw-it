# PROJECT KNOWLEDGE BASE

**Generated:** 2026-05-04
**Updated:** 2026-08-23
**Project:** PMW IT Service Portal (formerly "IT Onboarding Portal")

## OVERVIEW
React 19 + Vite 8 SPA with Azure AD MSAL authentication. Deployed on Vercel.
Forms are plain React over a small in-repo kit; SurveyJS was removed on
2026-08-23 (it took ~1.4MB of the bundle with it).

The UI is the SI CMMS shell — a branded nav column, a sticky bar, a dashboard of
stat cards over a canvas — minus everything maintenance-specific (work orders,
machines, priorities, SLA, roles). The sign-in screen is the PMW OSHE portal's
split poster/card layout, with an idle animation of its own.

## STRUCTURE
```
pmw-it/
├── src/
│   ├── components/
│   │   ├── form/             # the form kit: Field, Inputs, Choices,
│   │   │                     # RepeatRows, Wizard
│   │   ├── AppShell.jsx      # nav column + sticky bar + main; the one auth gate
│   │   ├── IdleAnimation.jsx # the sign-in screen's chip/packet animation
│   │   ├── Logo.jsx          # PMW mark (src/assets/logo-*.png)
│   │   ├── SessionDialog.jsx # what a timed-out session looks like being fixed
│   │   ├── SignInTransition.jsx # post-sign-in veil, fades into the dashboard
│   │   ├── SignatureDialog.jsx
│   │   └── ui/               # Icons, Button, Surfaces, StatCard, Badges
│   ├── features/
│   │   ├── semantic/         # ingest/ profile/ clean/ engine/ canvas/
│   │   │                     # intent/ suggest/ text/ export/ worker/
│   │   ├── devices/          # parse/ derive/ sharepoint/ stats/ ui/
│   │   ├── forms/            # the two forms' fields, validation and writes
│   │   ├── assets/           # scan/ draft/ handover/ people/ store/
│   │   │                     # sharepoint/ stats/ ui/
│   │   └── sharepoint/       # spClient, writePool, provision (shared plumbing)
│   ├── hooks/
│   │   ├── useRequests.js    # the one SharePoint read + row helpers + token
│   │   └── useSession.js     # session context + phases (no component here)
│   ├── pages/                # Homepage, LoginPage, DashboardPage, ListPage,
│   │                         # FormPage, AssetChecklistPage, DevicesPage,
│   │                         # AssetsPage + Scan/Batch/Detail/
│   │                         # Handover/People/Person
│   ├── styles/
│   │   ├── shell.css         # tokens, brand surface, shell, UI, dashboard
│   │   ├── auth.css          # sign-in layout + the idle animation
│   │   ├── devices.css       # the device list section
│   │   ├── assets.css        # the asset inventory section
│   │   ├── forms.css         # the form kit
│   │   └── semantic.css      # Semantic Analysis, incl. the chart series palette
│   ├── context/              # ThemeContext (dark/light), SessionContext (auto
│   │                         # re-sign-in + its dialog and entrance animation)
│   ├── services/             # sharePointService.js
│   ├── utils/                # timeout.js, initials.js, authErrors.js,
│   │                         # sessionKeys.js
│   ├── App.jsx               # Router setup
│   ├── main.jsx              # MSAL bootstrap + providers + stylesheet order
│   └── authConfig.js         # Azure AD + SharePoint scopes
├── public/                   # Static assets
├── vite.config.js
├── eslint.config.js
└── package.json
```

## ROUTES
| Path | Screen |
|------|--------|
| `/` | Redirects to `/dashboard` or `/login` |
| `/login` | Sign-in poster + card |
| `/dashboard` | Stat cards, charts, latest requests |
| `/requests` | The records table; filters live in the query string |
| `/list` | Legacy alias, redirects to `/requests` (keeps the query) |
| `/it-boarding-form` | HR/manager raises an onboarding or offboarding event; several employees per submission, `?edit=<id>` opens a record |
| `/asset-checklist` | The EMPLOYEE's own signed record of what they received or handed back — IN / OUT / INDIVIDUAL REQUEST, following the supplied reference form |
| `/devices` | Device list: fleet dashboard, register and scan-report import (`?view=`) |
| `/assets` | Asset inventory: what IT owns, its figures, and the deliveries still unsaved on this device (`?category=`, `?status=`, `?condition=`, `?location=`, `?unlabelled=1`) |
| `/assets/scan` | Purchase details, then the camera. Scans become a batch on this device — nothing reaches SharePoint here |
| `/assets/batch/:id` | Review a scanned delivery and save it |
| `/assets/:id` | One item in full, editable and removable, with who holds it and its handover history |
| `/assets/handover` | Pick a person, fill a basket by search or camera, hand it over |
| `/assets/people` | Everyone currently holding something, overdue first |
| `/assets/people/:email` | One person, everything they hold, and returning it |
| `/semantic-analysis` | Drop a Microsoft Forms export and land on a finished screen. It reads the file name for the subject, parks the form's bookkeeping columns, charts the rest, reads the written answers with a local model, sorts them into categories (internet, SAP, digitization, paperwork…) and charts those too. Tapping any mark filters the response list below — email, submitted, department, then every answer — and a response opens in full. Nothing is uploaded and nothing is saved: no SharePoint, no IndexedDB. Charts export as PNG, responses as CSV. Lazy route; `/data-studio` redirects here. |

## WHERE TO LOOK
| Task | Location |
|------|----------|
| Auth logic | `src/main.jsx` (MSAL init), `src/authConfig.js` |
| Auth gate for a page | `src/components/AppShell.jsx` — pages do not gate themselves |
| Timed-out session / auto sign-in | `src/context/SessionContext.jsx` |
| What counts as a dead session | `isInteractionRequired` in `src/utils/authErrors.js` |
| Any SharePoint token | `useSharePointToken()` in `src/hooks/useRequests.js` |
| Routes | `src/App.jsx` |
| Nav items | `NAV_ITEMS` in `src/components/AppShell.jsx` |
| Design tokens / layout | `src/styles/shell.css` |
| Sign-in screen | `src/pages/LoginPage.jsx`, `src/styles/auth.css` |
| SharePoint reads | `src/hooks/useRequests.js` |
| Device report parsing | `src/features/devices/parse/` |
| Device derived fields and risk | `src/features/devices/derive/` |
| Drives that belong to IT, not to the machine | `src/features/devices/derive/itMedia.js` |
| What each department asks of a machine | `src/features/devices/derive/persona.js` |
| Whether a machine suits the desk it is on | `src/features/devices/derive/deviceFit.js` |
| Office licensing, graphics and server dependency | `src/features/devices/derive/officeLicense.js`, `gpuClass.js`, `serverDependency.js` |
| A form's fields, branching and validation | `src/features/forms/` |
| The form controls themselves | `src/components/form/` |
| What the checklist writes to SharePoint | `src/features/forms/toChecklistItem.js` |
| Adding options to an existing choice column | `mergeChoices` in `src/features/sharepoint/provision.js` |
| Which values on a device page read red or green | `src/features/devices/fieldTone.js` |
| Any SharePoint list/column/view provisioning | `src/features/sharepoint/provision.js` |
| The SharePoint fetch wrapper and binary upload | `src/features/sharepoint/spClient.js` |
| Concurrency and retry for SharePoint writes | `src/features/sharepoint/writePool.js` |
| Which barcode on a box is the serial | `src/features/assets/scan/classifyCode.js` |
| Why a code held in frame is not counted twice | `src/features/assets/scan/scanSession.js` |
| The camera, its permissions and the decode loop | `src/features/assets/scan/useScanner.js` |
| Native vs ponyfill barcode decoding | `src/features/assets/scan/detector.js` |
| Reading the printed words off a label | `src/features/assets/scan/textReader.js`, `useTextScanner.js` |
| Which line of a label is which field | `src/features/assets/scan/classifyText.js` |
| Why a read value must agree twice, and what a scan may overwrite | `src/features/assets/scan/textScan.js` |
| Where the recognition engine is served from | `scripts/fetch-ocr.mjs`, `public/ocr/` |
| What makes two rows the same asset | `src/features/assets/identity.js` |
| Tracked-vs-bulk, and the category list | `src/features/assets/assetKinds.js` |
| The individual items inside a bulk line | `src/features/assets/units.js`, `ui/UnitPager.jsx` |
| What a bulk row may not hold, and where it goes instead | `units.js` — `PER_UNIT_ONLY`, `withUnitsSplitOut` |
| Why a saved photo needs the API and not the library | `src/features/assets/sharepoint/fileUrl.js`, `ui/useSharePointImage.js` |
| A delivery held offline | `src/features/assets/draft/batch.js`, `store/assetDb.js` |
| What a save writes, updates or refuses | `src/features/assets/sharepoint/planSave.js` |
| Owned vs out vs available | `src/features/assets/handover/availability.js` |
| What a handover writes or refuses | `src/features/assets/handover/planHandover.js` |
| What a return writes | `src/features/assets/handover/planReturn.js` |
| Finding a person in the directory | `src/features/assets/people/peopleSearch.js` |
| Handover SharePoint schema | `src/features/assets/sharepoint/handoverSchema.js` |
| The handover and return writes | `src/features/assets/sharepoint/writeHandover.js` |
| Asset SharePoint schema and views | `src/features/assets/sharepoint/assetSchema.js`, `assetViews.js` |
| Editing or removing one asset | `src/features/assets/sharepoint/updateAsset.js` |
| Device SharePoint schema | `src/features/devices/sharepoint/deviceSchema.js` |
| Device SharePoint list views | `src/features/devices/sharepoint/deviceViews.js` |
| Editing or removing one device row | `src/features/devices/sharepoint/updateDevice.js` |
| Removing several device rows at once | `deleteDevices` in `src/features/devices/sharepoint/updateDevice.js` |
| What the register has ticked | `src/features/devices/selection.js` |
| Device fleet statistics | `src/features/devices/stats/deviceStats.js` |
| Bar and column charts | `src/components/ui/Charts.jsx` (shared by both dashboards) |
| SharePoint writes | `src/services/sharePointService.js` |
| Theme | `src/context/ThemeContext.jsx`; toggle lives in the shell's bar |
| Spreadsheet parsing / header detection | `src/features/semantic/ingest/` |
| Column type inference and stats | `src/features/semantic/profile/` |
| The cleaning ops, proposals and apply | `src/features/semantic/clean/` |
| Columnar store, filter masks, aggregation | `src/features/semantic/engine/` |
| Chart tiles, theme, grid | `src/features/semantic/canvas/` |
| Starter chart suggestions | `src/features/semantic/suggest/suggestCharts.js` |
| PNG and CSV export | `src/features/semantic/export/exporters.js` |
| Malaysia time parsing and formatting | `src/utils/malaysiaTime.js` |
| Semantic Analysis state and worker lifetime | `src/features/semantic/SemanticContext.jsx` |
| Multi-select column detection and encoding | `profile/inferType.js` (`detectMultiSeparator`), `engine/dataset.js` (`encodeMulti`) |
| Splitting a written answer into issues | `src/features/semantic/text/splitIssues.js` |
| The analysis categories and their descriptions | `src/features/semantic/text/buckets.js` |
| The model, and where it is served from | `src/features/semantic/text/embed.js`, `scripts/fetch-model.mjs` |
| Text analysis pipeline and its worker | `text/analysis.js`, `worker/text.worker.js` |
| The user's corrections to the analysis | `src/features/semantic/text/overrides.js` |
| What the file NAME says the sheet is about | `src/features/semantic/intent/fileIntent.js` |
| Which columns are form bookkeeping | `src/features/semantic/intent/adminColumns.js` |
| The one decision taken per import | `src/features/semantic/intent/planAutopilot.js` |
| The card that discloses that decision | `src/features/semantic/intent/AutoBrief.jsx` |
| Decoding one stored cell into readable text | `src/features/semantic/engine/formatCell.js` |
| The responses behind the charts, and one in full | `engine/rows.js`, `canvas/ResponsePanel.jsx` |
| Who answered vs what they answered | `src/features/semantic/engine/responseFields.js` |
| The charts built out of the reading | `text/analysisTiles.js`, `text/chartAnalysis.js` |

## CONVENTIONS

**Page composition**: a screen renders `<AppShell title subtitle actions>` and
its own body. The bar, the nav, the theme toggle, sign-out and the sign-in gate
all belong to the shell — do not re-add per-page copies of them.

**Stylesheet order** (`src/main.jsx`): `index.css` → `App.css` → `styles/shell.css`
→ `styles/auth.css` → `styles/devices.css` → `styles/semantic.css`.
shell.css re-points the older `--bg` / `--surface` /
`--border` / `--text-*` tokens at the new palette, so it must load last.
`--accent` is deliberately left alone: it fills `.ms-button`, whose text colour
is `--bg`.

**Dashboard ↔ records**: every dashboard figure links into `/requests` with a
query string (`?type=`, `?entity=`, `?department=`, `?range=`, `?equipment=`).
Both screens read the same `useRequests()` fetch, so a card and the list it opens
cannot disagree.

**Navigation**: use `window.location.replace()` instead of React Router
`navigate()` *inside `useEffect`*. WHY: navigate causes a state update →
re-render → effect runs again → infinite loop. In event handlers `navigate()` is
correct and is what the shell and pages use.

**MSAL redirect handling**: always await `handleRedirectPromise()` before
rendering. Silent `no_token_request_cache_error` is normal on fresh load.

**Session guard** (`src/context/SessionContext.jsx`): a timed-out session signs
itself back in. Two rules keep it from disturbing anyone who is still signed in:

1. Nothing is on a clock — no idle timer, no expiry watcher. A recovery starts
   only on proof (an `InteractionRequiredAuthError`, or MSAL's silent renewal
   coming back `timed_out`), or when the account has vanished from the cache on
   a page that needs one. Network failures and SharePoint errors stay errors.
2. A recovery tries `ssoSilent` first. A live Azure AD session comes back in
   about a second and nobody leaves the page; only a second refusal escalates to
   `loginRedirect`, guarded by a one-shot `sessionStorage` flag so a sign-in that
   keeps failing cannot bounce the browser in a loop.

`/login` is exempt, and sign-out clears the stored login hint *before* the
redirect — leaving it would let the guard read a deliberate sign-out as a
timeout and undo it. Sign out through `useSession().signOut()` for that reason.

**`src/features/<name>/`** is where a section with more than a handful of modules
lives — `semantic/` and `devices/` both follow it. Layering inside a feature:
`parse/` knows nothing about the domain, `derive/` knows nothing about SharePoint,
`sharepoint/` imports no React. Each layer is testable without the one above it.
Semantic Analysis layers the same way — `ingest/` → `profile/` → `clean/` → `engine/`,
each a set of pure functions over plain data, with `canvas/` the only part that
touches React. That is why `engine/aggregate.js` carries more tests than the
whole canvas does: it is the part that can be wrong without looking wrong.

**An import decides everything at once, in `intent/planAutopilot.js`.** There is
no profile screen and no cleaning checklist to park the user on: the file name is
read for the subject, the form's bookkeeping columns are set to role `ignored`,
the starter charts are ranked with the title's keywords as a nudge, and the
written answers are read in the background — always, since reading them is what
this section is for. The reading then charts ITSELF, in `text/chartAnalysis.js`:
five derived columns go onto the grid and the category, theme and severity charts
go above the rest. A later correction re-scores the analysis but leaves those
charts alone until the user asks, so nothing is rebuilt under them mid-read.
The provider carries the plan out and takes no decisions of its own, so the
whole of the behaviour is testable without React — see `intent/*.test.js`.

Everything the plan decided is disclosed by `intent/AutoBrief.jsx` at the top of
the canvas, and every part of it is reversible from there in one click. That
card is not decoration: it is the reason an autopilot that reads header names
with a lexicon is safe to ship at all.

**The scan is run off a USB disk, and that disk is in the reports.**
`derive/itMedia.js` lists it by exact model (`WDC WD10 JPVX-60JC3T1`) and
`deriveStorage` keeps it out of the total, the drive count, the disk type and
therefore the risk score -- counted, it adds 932 GB the machine does not have,
turns an all-SSD laptop into "Mixed", and charges 10 points for a spinning disk
that is not in it. The drive is still reported, under `ignoredDrives`, so the
device page can say why the numbers are lower than the report. Match on the
exact model, never on "WDC" or on the size: a real 1 TB disk inside a desktop
has to keep counting. Rows already in SharePoint keep their old figures until
the reports are imported again.

**Every processor lands on one scale: the Intel generation it is contemporary
with** (`cpuGenerationRank`). A Ryzen badge does not carry its age -- a Ryzen 5
7530U and a Ryzen 5 7640U are both "7000" and are three years of architecture
apart -- so `deriveCpu` reads the architecture out of the model number and
places the Zen it finds against Intel: Zen 3 with 11th gen, Zen 4 with 13th,
Core Ultra 1 and 2 continuing the count at 14 and 15. Two rules do the work, and
both matter: on 2022-and-later MOBILE parts the THIRD digit is the architecture
(7*3*30U is Zen 3, 7*8*40U is Zen 4), and mobile parts before that run one
series behind their desktop namesakes (a 3500U is Zen+ where a desktop 3600 is
Zen 2). The digit rule is mobile-only -- applied to a desktop 7950X it would
read the 5 and call a Zen 4 chip Zen 5.

AMD is more than Ryzen, and the rest of it has to be recognised too. A
Threadripper carries no tier digit; a `PRO` sits where the tier digit would be;
the Zen-based Athlons (3000G, 3050U, the Gold and Silver laptop parts) are on
the scale at Zen; and everything AMD built before Zen -- `A8-7410`, `FX-8350`,
Phenom, Sempron, Turion, `Athlon II` -- is named in `OBSOLETE_FAMILIES`
alongside Pentium and Celeron. That last list is load-bearing: without it those
parts reach the RAM-type fallback and a DDR4 board alone calls a 2014 APU
Aging. Vendor detection reads the family names as well, so a report that says
`Athlon(tm) II X2 240` and never says "AMD" is not counted as `Other`.

**A machine is judged against the desk it sits on, not against one fleet-wide
bar.** `persona.js` maps a department to a workload profile — Engineering /
Technical / Media, Logistics / Operations / Desk, Executive / Field — and
`deviceFit.js` grades every machine against that profile as Critical, Needs
Attention, Moderate or Optimal, writing out the sentence behind each verdict.
This sits BESIDE `riskScore.js`, which is unchanged and still department-blind:
risk asks "is this machine dangerous?", fit asks "is it the right machine for
this person?", and 16 GB with no graphics card can pass one and fail the other.

The whole persona layer is computed on read (`enrichFit.js`, applied by
`deriveDevice` on import and `refixStored` on every SharePoint read) and NOTHING
of it is stored. Every ingredient it needs is already on the row, so changing
the memory floor for Engineering re-grades the fleet on the next page load with
no re-scan and no column migration.

The portability suggestion is a label, never a fault: a desktop in a field role
is tagged and counted, and does not on its own move a machine out of Moderate.

**A device page colours its values, and can be told not to.** `fieldTone.js`
holds the judgement -- red for what needs attention, green for what does not --
and only for fields with a settled right answer. A RAM discrepancy, a static IP
and a free memory slot stay the colour of the rest of the page on purpose:
colour that appears everywhere says nothing. The toggle sits in the page header
and is remembered per browser under `deviceValueTones`.

**Device report parsing keys off a known-label whitelist** (`parse/labels.js`). A
generic `^Word:` split reads `Total Slots: 2 | Used Slots: 2` and
`Y: | \\server\PMW\IT` as field names and moves those values out of the blocks they
belong to. An unknown label owns the lines beneath it, so a field the scan script
adds later surfaces in review rather than contaminating its predecessor.

**A hand-edited device field outranks the scan file.** The register lets the
three DERIVED fields be retyped (owner, department, device type) and records
which ones in `ManualFields`. `applyManualOverrides` in `syncDevices.js` then
holds those back on re-import — from the diff AND from the body, or updating
anything else would overwrite them as a side effect. Clearing a field is how
somebody hands it back to the scan file; without that it would stay frozen
against every future import for good.

**A tick in the register cannot reach off screen.** The register removes several
machines at once, and `selection.js` holds the one rule that makes that safe:
everything read out of the ticked set is scoped to the rows the filters are
SHOWING, and every write back is pruned to them too. Without it, ticking three
machines and then narrowing the search would leave "Remove 3" pointing at rows
nobody could see. It is deliberately not an effect — an effect would set state
during the render that caused it, and would still have to be right about the
same thing. A row with no `id` cannot be ticked at all, because it cannot be
removed either.

**The register freezes two columns, and one of them is a single cell.** Forty-odd
scan columns wide, scrolling to the end of a row would otherwise leave somebody
pressing Remove on a machine they can no longer name. The tick box and the
computer name share ONE `.dt-identity` cell rather than two pinned side by side:
two would need the second one's `left` offset to equal the exact rendered width
of the first, which a table decides for itself. Every frozen cell also carries
its own opaque background, and each row state repeats its colour on those cells
-- a sticky cell is painted while the row slides underneath it, so a transparent
one shows both at once and a highlight breaks at the frozen edge.

**`Total RAM` in a scan report is usable RAM, not installed RAM.** Windows subtracts
the integrated GPU's reserved share, so a 16 GB laptop reports 15 GB and an 8 GB one
reports 7 GB. Sum `RAM Slot Info` for the real figure; ranking on the reported one
puts a 16 GB machine below an 8 GB machine.

**A scanning session is an object, not a connection.** `/assets/scan` writes
nothing to SharePoint. Codes and photos go to IndexedDB (`assets/store/assetDb.js`)
as a BATCH, and only `/assets/batch/:id` reaches the network. That is what makes
a store room with no signal a non-issue rather than a feature, and it is why
there is no retry queue: a batch that has not been saved has no half-written
state to reason about. The price is that an unsaved batch is invisible to
everybody else, which is what the undismissable banner on `/assets` is for.

**An asset is identified by `AssetKey`, never by its name.** `assets/identity.js`
derives it: `serial:<make>|<serial>` for a tracked unit, `bulk:<category>|<make>|<model>`
for a bulk line, falling back to the sticker label and then to a `local:` key that
admits it will never match again. Re-scanning a laptop updates its row; a second
bag of the same mice ADDS to the quantity. Correcting a serial number re-derives
the key, because it changed which physical thing the row claims to be.

**Two kinds of asset, decided by the category.** `assetKinds.js` maps category to
tracked-or-bulk so the question is never answered twice or differently by two
people. Tracked rows are pinned to a quantity of 1 wherever they are touched —
twenty units cannot share one serial number.

**A serial, a part number, a MAC, a label, a condition and a status describe
ONE thing, so a bulk row does not hold them.** `PER_UNIT_ONLY` in `units.js` is
the list, and `withUnitsSplitOut` is the invariant: wherever a bulk row is
written — the edit screen through `planEdit`, a delivery through `planSave` —
those six come off the row and onto its items. A row carrying one serial for
twenty items is not a record of twenty items; it is a record of one, with
nineteen hidden behind it. A row saved before the rule has them MOVED onto item
1 rather than deleted, and `unitsOf` reads them there from the moment the rule
landed, so nothing looks lost in between.

**A handover can be signed for, both ways.** `SignatureField` on the handover
and person pages draws into the existing `SignatureDialog`; the PNG goes to the
asset photo library (`sharepoint/uploadSignature.js`) and the row keeps the
path in `IssueSignature` / `ReturnSignature`. RECOMMENDED, never required: a
signature that will not upload is reported and the handover is still recorded,
because the laptop changed hands either way. A blank means nobody signed and is
never written over an existing one.

**Rows that are one thing bought ten times can be put back together.**
`combine.js` (pure) plans it and `sharepoint/combineAssets.js` writes it: the
oldest row survives, every other row's serial, label, condition and status
becomes an ITEM on it, the quantity is the sum, and the other rows are removed
-- survivor written first, so a failure halfway leaves everything twice rather
than losing it. Picked on `/assets` behind "Combine rows". Refused while any of
them is out with somebody, because a handover names the row it came from.

**A bulk line knows its individual items.** One UNIT RECORD per physical thing
(`units.js`, paged one at a time by `ui/UnitPager.jsx` on `/assets/:id`, by the
arrows -- the sideways swipe was removed, it fought the text boxes on the card). They live as one JSON string in the `Units` column, SPARSE and with
blank fields dropped: only the units somebody has written on are stored, so a
box of twenty cables costs nothing until one of them is written on. The count
follows the row's quantity and lowering it only HIDES units — a quantity typed
wrong and corrected back must not take a serial number with it. The change log
gets a line per unit and field ("Unit 2 · Serial number"), never the JSON.
`planSave` carries the column across a re-scan by hand, because a draft has
none and the save writes every column; a re-scanned bulk box takes the NEXT
position (`mergeUnits`), never item 1's, because it is a different object.

**A scan may claim codes for an item, never a condition.** `PER_UNIT_CODES` is
the narrower list the review grid is allowed to move onto an item. "All new" on
a review grid is about the delivery, and writing it onto item 1 alone turns
twenty new cables into one new cable and nineteen nobody looked at — so the
condition field is not offered on a bulk draft at all.

**Anything that counts a per-item field counts ITEMS.** `perItem` /
`countPerItem` do the arithmetic once and everything else reads them: a box of
twenty with one Faulty is one faulty, not twenty and not one row. Items nobody
has spoken for are 'In stock' for a status and nothing at all for a condition —
a condition nobody recorded is not a condition. Search
(`assetFilters.haystack`) and the handover scanner both reach into the units,
or a tab could not be found by the serial on the tab in your hand.

**Which barcode is the serial is a guess, and is labelled as one.** A printed
label does not say. `scan/classifyCode.js` takes an explicit `S/N:` prefix as
fact, a colon-separated MAC as fact, a retail EAN/UPC as a PART number (every
identical monitor on the pallet carries the same one), and scores the rest by
shape. Everything it infers lands in `guessed` and renders as `guessed` in the
review grid — same contract as the device import's derived values, and the
reason a shape heuristic is safe to ship.

**A code seen on TWO boxes is the part number.** The surest evidence there is,
because a serial appears on one box and one box only. In ONE-item mode a repeat
arriving while a new box is being pooled is kept (`OUTCOMES.SHARED`) instead of
refused, and `classifyCodes` files it as the part number before anything is
scored — which is what lets the second of two identical tabs keep its own
serial. A repeat arriving on an EMPTY pool is still refused: that is the box
just confirmed being carried back into frame. Shape is the fallback, and
`partScore` reads what `serialScore` cannot — `#ABU` and `/A` are vendor part
suffixes no serial scheme prints. Where it still gets it backwards, the review
grid swaps both fields in one press (`swapSerialAndPart`).

**SharePoint plumbing is shared, not per-section.** `features/sharepoint/`
holds `spClient.js`, `writePool.js` and `provision.js`. `provisionSchema` takes
`{ lists, views }` and knows nothing about what is in them; `devices/sharepoint/
provisionLists.js` and `assets/sharepoint/provisionAssets.js` are both thin
declarations over it. Every SharePoint column rule listed below lives in that one
file now.

**`Quantity` is what the company OWNS and never moves when something is handed
out.** `QuantityOut` counts what is with people, and available is the
difference (`handover/availability.js`). A box of twenty with three out reads
20 / 3 / 17, not a box of seventeen. This means a return only ever moves the
derived figure, so a handover nobody recorded cannot silently change how much
the company believes it bought. Every row saved before handovers existed reads
`quantityOut` as 0, which is correct and needed no migration.

**The handover list is the truth; the register row carries a readable copy.**
`AssignedTo`, `AssignedOn`, `DueOn` and `Status` on an asset are copies of its
open handover so a row opened in SharePoint reads without a join — and they are
only ever written on a TRACKED row, because a box of cables can be with five
people at once and there is no honest single value. On a bulk row those stay
empty and `QuantityOut` carries the answer.

**A person is identified by email, never by name.** `peopleSearch.js` uses
SharePoint's own people picker (`clientPeoplePickerSearchUser`) rather than
Graph `/users`, which would need `User.ReadBasic.All` consented by an admin
before the feature worked at all. The picker answers with a JSON STRING inside
its JSON, and email lives in three different places depending on how the
account reached the directory — `normalisePerson` handles both.

**`toUpdateItem` and `toListItem` are not interchangeable.** `toListItem` writes
EVERY column, which is right when a whole record is being saved and catastrophic
when it is not: a handover setting `quantityOut` through it would blank the
serial number, supplier and photo of every item it touched. Partial writes go
through `toUpdateItem`, in both `assetSchema.js` and `handoverSchema.js`.

**The two forms are different things.** `/it-boarding-form` is HR or a manager
raising an onboarding or offboarding EVENT — a request, feeding `IT Request
Form`, which the dashboard and `/requests` are built on. `/asset-checklist` is
the EMPLOYEE signing for what they received or handed back, feeding
`Asset Checklist Form`. They overlap in subject and in nothing else. Neither
touches the asset register, which is IT's own.

**A form's fields and rules are data; the page only draws them.**
`features/forms/checklistForm.js` declares what each mode asks,
`validate.js` says what counts as complete, and `toChecklistItem.js` builds the
SharePoint row. All three are pure and tested, which is the point: "an OUT
checklist needs a signature" and "an individual request needs at least one
item" are exactly the rules that can be wrong without looking wrong.

**What the checklist READS and what it STORES are deliberately different.** The
form says IN / OUT / INDIVIDUAL REQUEST, as the reference form does; the list
keeps `In` / `Out` / `Individual Request`, which is what every checklist ever
signed already holds. The mapping is `FORM_MODES` and lives in one place.

**Choice columns are reconciled additively, never destructively.**
`mergeChoices` in `features/sharepoint/provision.js` adds declared options that
a column is missing and removes nothing. A SharePoint row holding a value no
longer in its column's list becomes unreadable in that list, so dropping
`pmw-ss` to tidy up would damage every record that used it. `provisionSchema`
does this for the sections it owns; `reconcileChoices` in
`sharePointService.js` does it for `IT Request Form`.

**SharePoint column creation**, verified against the tenant on 2026-08-21 while
provisioning the device lists. All three of these fail silently or confusingly
if you get them wrong:

1. **Create each field as its concrete type**, not the base `SP.Field`.
   `SP.Field` does not declare `Choices`, so a choice column sent that way
   fails with *"The property 'Choices' does not exist on type 'SP.Field'"*.
   Use `SP.FieldChoice`, `SP.FieldNumber`, `SP.FieldDateTime`,
   `SP.FieldMultiLineText`.
2. **The internal name comes from the `Title` a field is CREATED with.**
   `StaticName` in the creation body does not control it. Create
   `Title: 'Device Type'` and the field is addressable only as
   `Device_x0020_Type`; every item write of `DeviceType` then fails with
   *"The property 'DeviceType' does not exist on type 'SP.Data...ListItem'"*.
   Create under the internal name, then MERGE the display `Title` on
   afterwards. This is what produced the hand-encoded `Calling_x0020_Name`
   in `sharePointService.js`.
3. **Read `InternalName`, never `StaticName`**, when checking which columns
   already exist. The two can disagree, and a column where they disagree is
   precisely the broken one.
4. **REST-created columns join no view.** A freshly provisioned list shows
   nothing but its Title until view fields are set explicitly, which is what
   `deviceViews.js` is for.
5. **`ViewQuery` is only honoured in the creation body.** A default view is
   never created, so a filter or sort declared on one has to be MERGEd on
   afterwards or it is silently dropped. Address the built-in view through
   `/defaultView`, not `getByTitle('All Items')` — that title is English-only.

`ensureAssetColumns` in `src/services/sharePointService.js` still has bug 1:
it sends `Choices` with `__metadata: SP.Field`. Its lists predate the bug, so
nothing is broken today, but the same code on a fresh site would fail.

**SharePoint scopes**: use ROOT domain only, never site paths.
- ✅ `https://pmwgroupcom.sharepoint.com/AllSites.Write`
- ❌ `https://pmwgroupcom.sharepoint.com/sites/IThelpdesk/AllSites.Write`

**Icons**: `src/components/ui/Icons.jsx`, transcribed on the 24px stroke grid.
No icon package is installed — add a glyph there rather than a dependency.

## ANTI-PATTERNS (THIS PROJECT)
- Don't use `navigate()` in useEffect — causes infinite loops
- Don't patch a Semantic Analysis profile's `columns` without re-running
  `retopProfile`. `topMeasure` and `primaryTemporal` are derived from the
  roles, and a profile patched by hand goes on naming a column that has since
  been ignored — which is how the starter dashboard opened with
  "Time taken over Timestamp" on a survey where the autopilot had parked both.
- Don't let a tile with no x axis fall through `makeXResolver`'s missing-column
  guard. "No x asked for" (a KPI) and "the x column is gone" (a stale saved
  dashboard) are different answers: the first aggregates the whole dataset as
  one group, the second draws nothing. Conflating them made every starter
  dashboard's KPI row read a confident `0` over a full sheet.
- Don't have the autopilot DELETE a column it thinks is bookkeeping. It sets
  role `ignored` and lists what it did in `AutoBrief`, because the lexicon in
  `adminColumns.js` reads header names and will eventually be wrong about
  somebody's sheet. Being wrong must cost a click, not a re-import.
- Don't start the text analysis on every import that happens to have a prose
  column. The first run pulls a 23MB model, so `planAutopilot` only sets
  `autoAnalyse` when the file TITLE says the writing is the data (a pain-point
  or feedback survey). Every other sheet gets a button.
- Don't add SharePoint scopes to loginRequest — separate request required
- Don't call `acquireTokenSilent` / `acquireTokenPopup` from a page. Use
  `useSharePointToken()`, which routes through the session guard. The popup
  fallback this replaced is where the "it just stopped loading" reports came
  from: a popup opened from an expired timer is not a user gesture, so browsers
  block it and the page waits forever on a window nobody was shown.
- Don't widen `isInteractionRequired` to catch more errors. Signing someone back
  in who never lost their session is worse than the error they were going to
  see; ambiguous codes belong on the "not a timeout" side.
- Don't use `useNavigate` for auth redirects — use window.location
- Don't add `assetsInclude: ['**/*.html']` to vite.config.js. It matches
  index.html itself, so Vite stops treating it as the HTML entry and emits it as
  an asset — `npm run build` then produces a dist/index.html containing
  `export default "/assets/index-….html"` and no bundle at all.
- Don't create a SharePoint DateTime column with `DisplayFormat: 0` when the time
  matters — that is DateOnly and silently discards it. Device columns use `1`,
  confirmed round-tripping a real instant back out of the list.
- Don't give `.bar-fill` anything but `display: block` in `shell.css`. It is a
  `<span>` inside `.bar-track`, which is a plain block, so it is not blockified
  the way a flex or grid child would be — left inline it ignores width and
  height and every dashboard bar paints as an empty track.
- Don't create a SharePoint Note column without `RichText: false`; a rich-text Note
  wraps stored values in `<div>` markup and will not round-trip.
- Don't send `Choices` on a base `SP.Field`. A property exists only on the type
  that declares it, and the tenant answers "The property 'Choices' does not exist
  on type 'SP.Field'". Choice columns go out as `SP.FieldChoice`.
- Don't create a SharePoint column under its display name. SharePoint derives the
  internal name from the `Title` a field is created with — `StaticName` in the
  creation body does not control it — so "Device Type" becomes
  `Device_x0020_Type` and every item write then fails with "The property
  'DeviceType' does not exist". Create the column under its internal name and
  MERGE the display name onto it afterwards, as `provisionLists.js` does. The
  hand-encoded `Calling_x0020_Name` in `sharePointService.js` is the same trap,
  paid the other way.
- Don't add `hour12` beside `hourCycle` in `malaysiaTime.js` — per the Intl spec an
  explicit `hour12` nullifies `hourCycle` entirely. The 24-hour path pins `h23`, the
  AM/PM path pins `h12`, and neither passes `hour12`.
- Don't "fix" the `xlsx` dependency back to a registry version. It is pinned to a
  tarball URL on `https://cdn.sheetjs.com`, and that is deliberate: the npm registry
  copy is frozen at 0.18.5 by a known registry bug, and that version carries two
  open high-severity advisories (prototype pollution GHSA-4r6h-8v6p-xvw6, ReDoS
  GHSA-5pgg-2g8v-p4x9) with no registry fix. `npm install xlsx`, or accepting a
  tooling suggestion to "resolve" the URL, silently downgrades and brings both back
  — and `npm audit` will then say nothing, because the CDN version is not in the
  registry's advisory graph at all. Upgrade by changing the version in the URL. Note
  the install therefore reaches cdn.sheetjs.com at build time; the lockfile pins an
  integrity hash so the contents are verified, but if that dependency is ever
  unacceptable, SheetJS documents vendoring the tarball into the repo instead.
- Don't import the `echarts` umbrella. It pulls in every chart type and both
  renderers — about a megabyte, most of it for charts this app never draws — and
  it does so silently, because the code still works. Import the default export
  of `src/features/semantic/canvas/echartsCore.js`, which registers exactly
  the pieces used.
- Don't shift a date-only column when a `sourceZone` is asked for. Adding eight
  hours to a value that has no time of day moves it to the wrong day, and nothing
  on screen shows that it happened. `castType` keys this off `dateOnly` and
  refuses; the section no longer exposes a zone control, but the op still takes
  the parameter and still has to be safe when something passes one.
- Don't replace the grid in state without sending it to the worker. The
  worker KEEPS the parsed grid so a re-clean costs one small plan message
  however large the sheet is -- so a main thread that rebuilds the grid
  (which is exactly what adding the text analysis does) leaves the worker
  cleaning the PREVIOUS sheet. Nothing errors: the analysis dashboard
  simply renders six tiles all reporting `Column "Severity" is not in
  this dataset`. `worker/gridSync.js` decides when to carry it, by
  identity, and the provider keeps `workerGridRef` in step.
- Don't decode a cell by checking `column.dictionary` before
  `type === 'multi'`. A multi column carries a dictionary too, so the
  dictionary branch reads its FLAT option array as one code per row and
  prints a real label from the wrong row -- a confident wrong answer, in a
  CSV somebody has already emailed on. `engine/formatCell.js` is the one
  decoder; the row table, the record card and the CSV export all read it,
  so there is one place to get this right rather than three.
- Don't give the records panel its own query. It reads `maskFor` with no
  tile id, which is the same mask the tiles read, so the list cannot
  disagree with the chart above it. A table that disagrees with its chart
  is worse than no table.
- Don't reset the records panel's page from an effect. `npm run lint`
  fails it (`react-hooks/set-state-in-effect`), and it renders one frame
  of the empty page first. The page offset is CLAMPED during render
  instead.
- Don't let Escape inside the record card reach the window. The canvas
  clears the cross-filter on Escape, so closing the card would also throw
  away the selection the user drilled in from and return them to an
  unfiltered dashboard. The card calls `stopImmediatePropagation` on the
  native event.
- Don't build the analysis dashboard by running `suggestCharts` over the
  derived columns. The suggester ranks by SHAPE because on an unknown
  sheet shape is the only evidence; here every column's meaning is known,
  and guessing would sum severity instead of averaging it -- ranking a
  category twenty people mentioned mildly above one three people are
  furious about. `text/analysisTiles.js` names the six tiles outright.
- Don't filter the tile that originated a cross-filter selection. Click "HR" on a
  department bar chart and the other tiles narrow to HR — but that chart keeps
  all its bars, or it collapses to the one bar just clicked and deletes the
  context needed to click anything else. `maskFor` in `engine/filterMask.js`
  implements it, its cache keys on whether the selection *applies* rather than on
  the tile id, and the source tile dims its unselected marks as presentation only.
- Don't write an invisible character as a literal in source. A U+00A0 or a
  zero-width space does not survive being retyped or pasted, so a test that
  claims to cover it silently becomes a no-op, and a diff cannot show the
  difference. Use ` `, `​`, `﻿`. This is not hypothetical here:
  one such swap voided a whole test in `inferType.test.js` and was only caught
  by byte-diffing the file.
- Don't let the analysis model come from anywhere but this app's own origin.
  `embed.js` sets `allowLocalModels = true`, `allowRemoteModels = false` and
  `localModelPath`. The tab tells the user nothing leaves their machine;
  those three lines are what make it true, and the first is a trap —
  `allowLocalModels` defaults to FALSE in the browser, so switching remote
  models off alone disables both and the pipeline refuses to start.
  `embed.contract.test.js` pins all of it by reading the file as text.
- Don't set `wasmPaths`. The runtime reaches its `.wasm` through
  `new URL(…, import.meta.url)`, which the bundler already rewrites to a
  hashed asset on this origin — so the default keeps the promise, and an
  override made the build ship the same 22.5MB file twice. If it ever has to
  be set, name the ONE file rather than a directory prefix: a prefix makes
  the runtime treat its loader as external too and fetch it with a dynamic
  `import()`, which Vite's dev server refuses to do for anything in `public/`,
  so the prefix form works in a build and fails in dev. The file wanted is
  `ort-wasm-simd-threaded.asyncify.wasm` — transformers.js imports
  `onnxruntime-web/webgpu` and that is what its bundled loader names. Grep
  the bundle for `ort-wasm` if it changes; guessing costs an afternoon.
- Don't point runtime code at a path under `public/` that only exists because
  somebody copied it there. `public/models` is gitignored and recreated by
  `scripts/fetch-model.mjs` (and `public/ocr` by `scripts/fetch-ocr.mjs`);
  anything else gitignored is not recreated by
  anything. A missing file does not 404 — the SPA rewrite answers with
  `index.html`, so the browser reports
  "expected magic word 00 61 73 6d, found 3c 21 64 6f", which is `<!do`.
  That cost a live bug; `embed.contract.test.js` is the guard.
- Don't re-embed on a settings change. The threshold and granularity controls
  re-score against cached fragment vectors, and bucket vectors are cached by
  prompt TEXT so renaming a bucket is free. Measured on the real survey:
  2143ms before the cache, 172ms after, against a 300ms budget for a control
  the user drags. Nothing about a regression here looks wrong in the code.
- Don't cluster by re-summing member distances on every merge. That is
  O(n^3) and took 3–10 seconds on the real survey's 134 fragments.
  `cluster.js` builds the distance matrix once and updates it with the
  Lance-Williams rule; `cluster.test.js` pins the cost.
- Don't score the priority list on how many issues were raised. It counts
  distinct PEOPLE, scaled by severity: one person writing five furious
  sentences must not outrank five people writing one each. `rankIssues.js`
  carries the test that pins it.
- Don't drop a fragment just because it starts with "no". "no issue from IT"
  is a non-answer; "No proper system exists for tracking approvals" is a
  report. `boilerplate.js` requires BOTH a leading non-answer word and a
  short body, and the pair has its own test.
- Don't gate free-text detection on the profiler's `text` verdict. The
  profiler calls any column with 50 or fewer distinct values categorical, so
  on a 42-response survey the written-answer column is never `text` and the
  whole feature would be invisible on exactly the files it was built for.
  `detectTextColumns.js` tests shape — length, fill, uniqueness — instead.
- Don't put `multi` handling after the `column.dictionary` branch in
  `aggregate.js` or `filterMask.js`. A multi column carries a dictionary too,
  so the dictionary branch catches it first and every option collapses into
  one meaningless category.
- Don't derive `dataset.rowCount` from the first column's `values.length`. A
  multi column's `values` is the flat option array and is longer than the
  grid, so every mask allocated from it would be the wrong size.
- Don't export a helper next to a component from the same file — it drops the
  file out of Fast Refresh (and eslint fails the build). `initialsOf` lives in
  `src/utils/initials.js` for exactly this reason.

- Don't let a photo failure take its row down with it. In `assets/sharepoint/
  saveBatch.js` the upload is caught per item and reported in `photoFailures`;
  the row saves without it. Losing the serial number of a laptop because the
  camera produced an odd JPEG is the wrong trade.
- Don't call `formatMYT` on a value that might not be an instant. It throws
  `RangeError: Invalid time value`, and one undated row would take a whole
  delivery's save down with it. `assetSchema.js` guards it in `readableMYT`.
- Don't block a labelled machine from being re-scanned. The tag-uniqueness check
  in `draftIssues` finds the row's OWN entry in the register, so it compares
  `owner.assetKey` against the draft's key first — without that, a labelled
  asset becomes the one kind that can never be updated.
- Don't drop the in-batch tag check and rely on the register one. Two rows in a
  single delivery can both claim `PMW-0142`, and neither is in SharePoint yet.
  `planSave` carries `claimedTags` for exactly this.
- Don't let one blocked row fail a whole delivery. A duplicate sticker label is
  one row's problem; `planSave` returns `blocked` beside `inserts`/`updates` and
  the rest still save.
- Don't re-save the rows that already landed. `remainingDrafts` leaves only what
  failed on the batch — the register would survive a re-write (the key upserts)
  but the change log would fill with phantom edits.
- Don't import the barcode ponyfill statically. `scan/detector.js` reaches it
  through `await import('barcode-detector/pure')` only when the browser has no
  native `BarcodeDetector`, which keeps 43KB of WebAssembly loader off every
  Android phone that never needs it.
- Don't set a font-size under 16px on an input in `assets.css`. iOS Safari zooms
  the whole page in on focus for anything smaller, and a zoomed viewfinder
  cannot be aimed.
- Don't write a ref during render (`handlerRef.current = onCodes`) — eslint's
  `react-hooks/refs` fails the build. `useScanner` assigns it in an effect, and
  keeps the handler in a ref at all so that a scan changing the session does not
  restart the decode loop and drop frames exactly when it is busiest.
- Don't clear per-row state from an effect on an id change. `AssetDetailPage`
  adjusts it during render against a `shownId`, because an effect paints one
  frame of the new item wearing the previous item's unsaved edits, and
  `react-hooks/set-state-in-effect` fails the build besides.
- Don't guess the photo library's folder path. A library titled "IT Asset Photos"
  does not reliably live at `/sites/…/IT Asset Photos`; `uploadPhoto.js` asks for
  `RootFolder/ServerRelativeUrl` once per save instead of 404ing on every upload.
- Don't send a Blob through `spFetch` — it JSON-stringifies the body and uploads
  the string "[object Blob]". Binary goes through `spUpload`.
- Don't put a stored photo path straight into an `<img src>`. What is saved is
  SERVER-RELATIVE (`/sites/IThelpdesk/IT Asset Photos/x.jpg`), so it resolves
  against the PORTAL's origin and asks Vercel for a file only SharePoint has —
  which is why no photograph in the register was visible until 2026-08-24. Nor
  is prefixing the host enough: the library path sends no CORS headers, so a
  cross-origin fetch fails before there is a status code, and it takes cookies
  rather than the app's token. Read the bytes through
  `/_api/web/GetFileByServerRelativeUrl('…')/$value` (`fileApiPath`), which
  does both, and keep `absoluteFileUrl` for the human-facing "open in
  SharePoint" link. Verified against the tenant on 2026-08-24: the library path
  answers `Failed to fetch`, the API path answers 401 for a bad token.

- Don't add a nav item without importing its icon. `NAV_ITEMS` in `AppShell.jsx`
  references the glyph by identifier, so a missing import is a ReferenceError
  inside the shell — which blanks EVERY page, passes `npm run build` (Vite does
  not type-check), and passes every test. Only opening the app catches it.
- Don't check a handover line's availability without first adding up the other
  lines of the same basket. Two lines of three each pass individually against a
  stock of five and hand out six; `coalesceLines` sums before the check, and
  `lineRefusal` counts sibling lines for the same reason.
- Don't clamp a return to what is out. "Two came back" and "three came back" is
  a real disagreement about what happened, and rounding it hides a miscount.
  `planReturn` refuses with the figure.
- Don't write two returns of the same box as two independent register updates.
  They have to accumulate against one row (`outByAsset` in `planReturn`), or the
  second write overwrites the first's arithmetic and half the return is lost.
- Don't clear an asset's assignment fields on a partial return. A bulk row with
  two of five still out must keep reading as partly out; only `fullyBack` clears.
- Don't write the register copies before the handover rows. The handover list is
  the truth, so a failed register update leaves an item's copied fields stale —
  recoverable. The reverse loses the record. `commitHandover` writes handovers
  first and only updates the rows whose handover actually landed.
- Don't trust the register as the screen had it when handing over. `commitHandover`
  re-reads it immediately before planning, which is what makes two people issuing
  the same laptop from two phones refuse the second one rather than both succeed.
- Don't put the overdue view on the register list. A box of cables held by three
  people has no single due date of its own, so the answer only exists per
  handover — the view lives on `IT Asset Handovers`.

- Don't parse a date-only input with `new Date('2026-09-01')`. That is UTC
  midnight, which in Malaysia is the previous day at 8am — so a date somebody
  picked stores as the day before. `parseDay` / `parseFormDate` build it at
  LOCAL noon instead.
- Don't remove an option from a SharePoint choice column, however stale it
  looks. Rows already holding that value become unreadable in their own list.
  `mergeChoices` only ever adds.
- Don't fetch a column's `Choices` per column. `existingFields` selects
  `InternalName,Title,Choices` in the ONE request it already makes; asking
  separately cost a round trip per choice column on every provisioning run,
  which `provisionLists.test.js` catches by asserting that an already-correct
  column costs nothing.
- Don't send an empty multi-choice list to SharePoint — it is rejected.
  `employeeToItem` omits the column instead.
- Don't reuse `toListItem` for a partial update; see the asset-register entry
  above. The same split exists in `features/forms/`.
- Don't clear a form's values when a submit fails. Retyping ten employees'
  details because the network blinked is the worst thing a form can do; both
  pages leave the answers on screen and offer Retry.
- Don't validate a whole multi-step form on step one. `validateChecklist` takes
  a `step`, because marking step two's fields red before somebody has reached
  them is hostile and tells them nothing they can act on yet.

## COMMANDS
```bash
npm run dev      # Start dev server on port 5173
npm run build    # Build for production (outputs to dist/)
npm run lint     # Run ESLint
npm run preview  # Preview production build
```

## NOTES
- Vite port 5173 is for local dev; Vercel ignores this
- MSAL handles Azure AD login flow + token caching
- Forms are plain React over `src/components/form/`
- `npm run lint` reports ONE remaining error, in `ThemeContext.jsx`
  (a non-component export breaking Fast Refresh). The FormPage,
  AssetChecklistPage and SignatureDialog errors are gone — they were SurveyJS
  model mutation inside hooks and unused imports, cleared by the rewrite.

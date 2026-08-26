# Delivery counting, missing paperwork, and the scanner — design

Date: 2026-08-26
Branch: `claude/inventory-delivery-scanner-fixes-1c27e4`
Status: approved by the user, not yet implemented.

## Why

Five complaints from the person filling the register, in their words:

1. A delivery of ten monitors was recorded as ten distinct items instead of one
   line of ten.
2. That delivery predates this feature. There is no DO number, no serial
   numbers and no photographs, and no way to say so or to finish it later.
3. On a phone the scan window sits in the page, so it has to be scrolled to.
4. The camera reads printed writing but writes it straight into the form. It
   should MARK what it read and let the person accept or discard each one.
5. Barcodes are frequently never decoded at all.

## What is already true

Reading these before changing anything saves re-deriving them:

- `assetKinds.js` maps a category to TRACKED or BULK. `Monitor` is TRACKED, so
  ten monitors are legitimately ten rows. `Tab` is deliberately BULK and is the
  precedent this design follows.
- A TRACKED row is pinned to `quantity: 1` in `setDraftField`, because twenty
  units cannot share one serial number. **This invariant is kept, not broken.**
- A BULK line already carries one UNIT RECORD per physical thing (`units.js`,
  paged by `ui/UnitPager.jsx`), sparse, stored as JSON in the `Units` column.
  Per-unit serial/part/MAC/label/condition live there. This is the mechanism
  section 1 leans on entirely.
- `withUnitsSplitOut` moves `PER_UNIT_ONLY` fields off a bulk row onto item 1
  wherever a bulk row is written (`planEdit`, `planSave`).
- A row with no serial and no label gets a `local:` key (`identity.js`) that can
  never match again — which is the second half of why ten rows stayed ten.
- `.as-sheet` (`assets.css:985`) is already `position: fixed`. The two IN-FLOW
  viewfinders are `AssetScanPage.jsx:247` and `AssetHandoverPage.jsx:308`
  (`.as-viewfinder`, `assets.css:316`). Those are the scrolling complaint.
- `TextScanSheet.jsx:38` auto-applies the scan on `SCAN_STATE.DONE`. That
  effect is what section 4 removes.
- `useScanner.js` decodes on a fixed 120ms timer, hands the raw `<video>`
  element to the decoder, and requests no focus mode. Section 5 is about that.
- OCR and barcode engines are both served from this origin (`public/ocr`,
  `barcode-detector/pure` imported dynamically). Do not reintroduce a CDN.

## 1. Quantity on every row

**Behaviour.** Every draft row shows a "How many?" box, not only bulk ones.
Entering a number above 1 on a tracked-category row turns the row into a
by-quantity line and says so in plain words under the box: "Counted as 10 —
each one keeps its own serial underneath." Returning it to 1 returns the row to
a single tracked unit unless the person set *Counted as* by hand.

**Mechanism.** In `setDraftField`:

- `quantity` is settable on any row. When the parsed value is `> 1` and
  `trackingMode` is TRACKED, the row flips to BULK and `trackingMode` joins
  `manualFields` (the flip is the person's decision, so a later category change
  must not undo it).
- The existing "a TRACKED row is quantity 1" clamp stays, but is applied only
  when the tracking mode is TRACKED *after* the flip above — so it can no longer
  silently discard a typed quantity.
- Lowering the quantity back to 1 does NOT flip back automatically. Only the
  *Counted as* dropdown does that, because lowering a count must never take a
  unit's serial with it (the existing rule that lowering quantity only HIDES
  units).

`DraftCard.jsx` shows the Quantity field unconditionally and moves the existing
"the serial below belongs to this one box" note so it appears whenever the row
is counted by quantity.

`AssetDetailPage.jsx` gets the same treatment on the saved row, so a line
recorded as one monitor can be corrected to ten later.

**Tests** (`draftAsset.test.js`): quantity 10 on a Monitor row yields
`trackingMode: BULK`, `quantity: 10`, `manualFields` containing `trackingMode`;
quantity 1 leaves a tracked row tracked; a category change after the flip does
not revert it; a serial already typed survives the flip and lands on item 1
through the existing `withUnitsSplitOut` path (assert via `coalesce` in
`planSave.test.js`).

## 2. "Details to follow"

**Behaviour.** A switch on the purchase step of `/assets/scan`: "An older
delivery I'm entering late — the paperwork is missing." When on:

- the non-blocking "no serial or label, so re-scanning this later will add a
  second row" issue is not raised. Blocking issues (duplicate sticker label, in
  the batch or in the register) are unaffected.
- every row in the delivery saves with `DetailsPending` true.
- `/assets` gains a "Needs details" filter (`?pending=1`) and a count beside the
  existing filters.
- `/assets/:id` shows a banner naming what is missing (serial, label, photo, DO
  number — whichever are blank) and a "Details are complete" button that clears
  the flag.

A **DO number** field is added beside the PO number, on the delivery and on the
row, because it is one of the things being backfilled.

**Mechanism.**

- `batch.js` `newPurchase()` gains `doNumber: ''` and `detailsPending: false`.
- `resolveDraft` inherits both onto each row, using the same
  `undefined`-means-inherit rule already there.
- `draftAsset.js` `newDraft` gains `doNumber: undefined`, `detailsPending:
  undefined`. `draftIssues` takes `detailsPending` from the resolved draft and
  skips the `hasStableIdentity` warning when it is true.
- `assetSchema.js` `ASSET_COLUMNS` gains `text('DoNumber', 'DO Number')` and
  `choice('DetailsPending', 'Needs Details', ['Yes','No'])` — a choice rather
  than a Yes/No column, to match how every other flag in this schema is
  declared. `BATCH_COLUMNS` gains `DoNumber` and `DetailsPending`. Provisioning
  picks both up automatically through `provisionSchema`.
- Add `doNumber` to `TRACKED_FIELDS` so a corrected DO number appears in the
  change log. Do NOT add `detailsPending` — it churns.
- `assetFilters.js` `filterAssets` gains a `pending` filter; `haystack` gains
  `doNumber`.
- `AssetsPage.jsx` renders the filter chip and count.
- `AssetDetailPage.jsx` renders the banner and the clear button.

**Tests**: `draftAsset.test.js` (warning suppressed only when pending, blocking
issues survive), `batch.test.js` (inherit / per-row override of both new
fields), `assetFilters.test.js` (the `pending` filter), `assetSchema.test.js`
(both columns round-trip).

## 3. The camera covers the screen

**Behaviour.** On viewports under 640px, opening any camera fills the screen:
picture edge to edge, controls and found-values list over it, close button top
right, the page behind it frozen. Desktop keeps today's centred panel.

**Mechanism.**

- `.as-sheet` on small screens: `height: 100dvh` (not `vh` — the mobile browser
  chrome makes `vh` overflow), `padding: 0`, `.as-sheet-inner` at
  `border-radius: 0; max-width: none; height: 100%`, the video growing to fill
  and the controls sitting over it.
- Body scroll is locked while a sheet is open. One shared hook, `useScrollLock`,
  so all four camera surfaces behave the same.
- `AssetScanPage.jsx` and `AssetHandoverPage.jsx` stop rendering
  `.as-viewfinder` in the page flow and render through the same overlay.
- Respect `env(safe-area-inset-*)` so the close button is not under a notch and
  the buttons are not under the home indicator.
- Keep the existing rule: no input font-size under 16px, or iOS zooms the page
  and the viewfinder cannot be aimed.

## 4. The camera reads writing, and asks first

**Behaviour.** The list under the picture builds as the camera reads. Each
settled value is a row: the field it thinks it is, the value, a tick and a
cross. Nothing enters the form until the tick. A crossed-out value does not come
back. A "Take all" button applies everything still listed. Writing it recognised
but could not name appears at the bottom as "Also on the label", each with a
small menu to file it into a field by hand. The sheet no longer closes itself.

**Mechanism.**

- Remove the auto-apply effect at `TextScanSheet.jsx:38`.
- `textScan.js` gains, on the scan object, `rejected: []` — values the person
  crossed out. `recordReading` skips a value present in `rejected` so the next
  pass does not re-offer it. Two pure helpers: `rejectValue(scan, field)` and
  `candidates(scan)` returning `[{ field, value, guessed }]`.
- The reading loop keeps running after values settle instead of stopping on
  `isComplete`, so more of the label can be picked up while the person decides.
  `MAX_PASSES` still ends it; `finish` still stops it early.
- Accepting one value calls the existing `setDraftField` contract (marks it set
  by hand, so a later scan cannot overwrite it) — the same thing the held-back
  "Use this instead" button already does in `DraftCard.jsx`.
- The existing `heldBack` list stays: it answers a different question (the scan
  read something over a value you had typed).

**Tests** (`textScan.test.js`): a rejected value is not re-offered on a later
pass; `candidates` reports field, value and guessed-ness; accepting one value
leaves the others pending; `applyScannedFields` is unchanged for the take-all
path.

## 5. Barcodes that are never decoded

Ordered by expected effect.

1. **Decode a crop, not the whole frame.** `detector.js` gains
   `frameToCanvas(video, region)` drawing the aiming box's area of the frame at
   the camera's native resolution. The loop alternates: crop, then whole frame.
   A small sticker currently occupies too few pixels to resolve; this is the
   single biggest cause.
2. **Sequential loop.** Replace the fixed 120ms `setTimeout` in `useScanner.js`
   with read-wait-read, as `useTextScanner.js` already does. On iOS the
   WebAssembly decoder takes longer than the interval and requests queue.
3. **Ask for focus and resolution.** `useCamera.js` and `useScanner.js`:
   `width/height ideal 1920×1080`, and `advanced: [{ focusMode: 'continuous' }]`.
   Both must stay `ideal`, never `exact` — a laptop with one fixed camera must
   still work.
4. **Torch.** Where `track.getCapabilities().torch` is true, a button over the
   picture calling `applyConstraints({ advanced: [{ torch }] })`.
5. **Zoom.** Where `capabilities.zoom` exists, a slider over the picture.
6. **Tap to focus.** Where `pointsOfInterest` is supported, tapping the picture
   sets it.
7. **Say something after silence.** After roughly eight seconds with nothing
   decoded: "Fill the box with the barcode, or turn the light on."
8. **The aiming box must mean something.** The video is `object-fit: cover`, so
   what is on screen is a crop of what the decoder sees. The crop in (1) is
   computed in the video's own coordinates from the rendered box, so the two
   agree.

**Tests**: `detector.test.js` (new) for the crop geometry — a reticle inset
expressed in rendered pixels maps to the right rectangle of a differently-sized
source under `object-fit: cover`, including when the aspect ratios differ.
Camera capability code stays thin and untested, matching the reasoning already
recorded for `useScanner`.

## Order of work

Each step is independently shippable and independently useful.

1. Section 1 — quantity on every row. Fixes the complaint that prompted this.
2. Section 2 — details to follow, plus the DO number.
3. Section 5 — barcode decoding. Highest daily friction of the remaining three.
4. Section 3 — full-screen camera.
5. Section 4 — accept/discard for read writing.

## Out of scope

- No change to how a category decides tracked-or-bulk. `Monitor` stays TRACKED;
  section 1 makes the per-row override reachable instead.
- No retry queue, no background sync. A batch stays on the device until saved.
- No new OCR or barcode engine, and no CDN.

## Resuming this work

Nothing below section 1 has been written yet. To continue in a fresh session:

    Read docs/superpowers/specs/2026-08-26-delivery-counting-and-scanner-design.md
    and implement section <n>, TDD, following AGENTS.md.

Check `git log --oneline` on this branch first — each section lands as its own
commit named after it.

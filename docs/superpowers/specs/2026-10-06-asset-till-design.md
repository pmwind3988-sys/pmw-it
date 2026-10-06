# Asset till — design

**Date:** 2026-10-06
**Status:** approved (mockup clicked through and signed off on 2026-10-06)
**Mockup:** Claude Design canvas "Asset Till", https://claude.ai/artifact/6PVmbdCq7hZkHYWMC317Du

## 1. Why

Registering a delivery takes four screens: purchase details → scan mode → camera →
a separate review page → save. Handing things to somebody is one long form. Both
jobs are, in practice, a cashier's job: scan things onto a list, check the list,
finish. The till makes them that, on one screen, on a phone or at a desk.

## 2. What it is

One page, `/assets/till`, with three modes in a tab strip at the top:

| Mode | Query | Receipt is | Checkout writes |
|------|-------|------------|-----------------|
| Stock in | `?mode=in` (default) | a delivery batch | `saveBatchToSharePoint` |
| Hand out | `?mode=out` | a handover basket | `commitHandover` |
| Take back | `?mode=back` | a list of returns | `commitReturn` |

Layout, top to bottom on a phone; left / right at ≥ 1024px:

1. **Scan area.** Camera viewfinder (phone: on by default; desktop: off, "Use
   webcam" button), and under it ONE text box that is always the answer to "scan
   or type". A USB scanner gun types into it and presses Enter, so it works with
   no setup. A **Can't scan?** button sits beside it and turns yellow after a
   failed or quiet camera.
2. **Context.** Stock in: a folded "Delivery" line (supplier, DO, PO, details
   to follow) that can be filled in at any time before saving. Hand out: who is
   getting it (`PersonPicker`), Loan / Keep, and a due date (1 week, 2 weeks,
   1 month, no date). Take back: the condition everything is coming back in.
3. **Receipt.** One line per thing, newest at the bottom. Bulk lines have − / +.
   A line that cannot go through says why, in place, the moment it is scanned.
4. **Footer.** The total, a note, and the one checkout button.

Desktop adds **quick keys** (the bulk lines most often handed out or delivered,
with stock left) and, in Take back, **out right now** (open handovers, one click
to add). Both are hidden on a phone, where the space belongs to the camera.

After checkout: a **done receipt** with a "Next delivery / person / return"
button that resets that mode only.

## 3. What a scan means

All of this is pure and lives in `src/features/assets/till/`, tested without React.

### Stock in (`stockIn.js`)

A scanned code, normalised with `normaliseCode`, is tried in this order:

1. **Already on this receipt.** Matches a draft's serial, tag, part number or
   extra codes. A bulk line counts up by one; a tracked line is a duplicate (the
   same box scanned twice) and nothing changes.
2. **Known to the register** (`matchRegister`, own serial / tag / part / extra
   codes, whole rows only):
   - by serial or tag → a draft carrying that row's identity, category, make and
     model. Saving it updates the row, as a re-scan does today.
   - by part number on a BULK row → a draft of that model, quantity 1. `planSave`
     adds the quantity onto the row.
   - by part number on a TRACKED row → another unit of that model: a draft with
     category, make, model and part number, and NO serial. It shows as "serial
     still to record" and `draftIssues` keeps it from saving until one is added.
3. **Unknown** → `draftFromCodes` as today. Its category is unconfirmed
   (`needsKind`): the line shows category chips and a model box, and checkout is
   held until each one is answered.

The receipt is a real batch (`draft/batch.js`) and is written to IndexedDB on
every change, so a phone that dies or loses signal loses nothing. The page keeps
the id of its open batch in localStorage (`tillBatchId`) and picks it back up on
the next visit. The "unsaved deliveries" banner on `/assets` lists it like any
other batch.

### Hand out (`handOut.js`)

The receipt is a basket (`handover/basket.js`). A scan goes through
`findScanTarget` exactly as the handover page does: a unit off a bulk row, or a
whole row. Every line carries `lineRefusal`. A code nothing in the register
carries becomes a refused line of its own ("Not in the register — stock it in
first"), held beside the basket, never in it. Refused lines are left off at
checkout and the footer says so.

### Take back (`takeBack.js`)

Needs the open handovers (`useHandovers`). A scan resolves to an asset (and unit)
with `findScanTarget`, then to the open handover(s) holding it:

- exactly one → a return line for that handover, one item (a tracked line or a
  unit line), or one off a bulk line, capped at what that person still has;
- several people hold the same bulk line → the line asks which one, as chips;
- nobody → refused: "Not out with anyone — it is in stock".

Checkout turns the lines into `[{ handoverId, quantity, condition }]` for
`commitReturn`. `planReturn` refuses over-returns; the till never clamps.

### Type to find (`tillSearch.js`)

The same box searches as you type (two characters or more), per mode:

- Stock in: models already in the register (category + make + model, once each).
  Picking one adds a draft of that model (bulk: quantity 1; tracked: no serial yet).
- Hand out: register rows with something available (`matchesQuery`, which already
  reaches into unit serials and tags).
- Take back: open handovers, matched on the item AND on who has it.

Enter always uses exactly what was typed, as a scan. A scanner gun never waits
for the list.

### Can't scan?

1. **Read the printed label** — the existing `TextScanSheet`. Stock in: the values
   go onto ONE new draft through `applyScannedFields(..., { byHand: true })`.
   Hand out / Take back: the first of serial, tag or part number is fed through the
   scan path.
2. **Type it** — focuses the box.
3. **Nothing on it** (Stock in only) — category, model, quantity. A tracked thing
   gets the next free label (`nextTag`: the highest `PMW-NNNN` in the register and
   on this receipt, plus one, four digits) so it can be found again. Bulk things
   get no label; their identity is the model, as it already is.

## 4. Checkout

- **Stock in.** Held while any line needs a category or `draftIssues` says it is
  blocked; each such line shows its message. Saves with
  `saveBatchToSharePoint`. Everything saved → the batch is deleted locally and the
  done receipt shows. Anything left (`remainingDrafts`) → the batch keeps only
  those rows and the page says how many, with a link to `/assets/batch/:id` for
  the full review grid. Photos are not taken on the till; the review page still
  takes them.
- **Hand out.** Needs a person. Opens the signing step: `SignatureField`, "Confirm
  handover", and "Hand over without a signature" — recommended, never required, as
  today. Writes with `commitHandover` (which re-reads the register itself, so two
  phones cannot issue the same laptop).
- **Take back.** Writes with `commitReturn`, no signature on the till (the
  person page still takes one).

Write phases show on the checkout button with the existing progress labels.
A failure leaves the receipt on screen with Retry; nothing is cleared.

## 5. What stays

`/assets/scan`, `/assets/batch/:id`, `/assets/handover` and the person page keep
working. The batch page is where a held delivery is finished in detail. The nav
gains **Till**, and `/assets` and `/assets/people` gain an **Open till** button.
Retiring the old scan and handover screens is a later decision, once the till has
been used on real deliveries.

## 6. Out of scope

Photos on the till, return signatures on the till, keyboard shortcuts beyond
Enter, printing label stickers.

## 7. Files

```
src/features/assets/till/
  stockIn.js        + test   scan → batch, register match, needsKind, nextTag, noCodeDraft
  handOut.js        + test   scan → basket line or refusal, due-date choices
  takeBack.js       + test   scan → return line, holder choice, toReturns
  tillSearch.js     + test   type-to-find per mode
  quickKeys.js      + test   which bulk lines get a key
  ui/TillCamera.jsx          inline viewfinder over useScanner
  ui/TillReceipt.jsx         the receipt lines for all three modes
  ui/CantScanSheet.jsx       the three ways round a stubborn code
  ui/NoCodeSheet.jsx
  ui/SignStep.jsx
  ui/DoneReceipt.jsx
src/pages/AssetTillPage.jsx  orchestration only
src/styles/till.css          loaded after assets.css
```

## 8. Testing

Unit tests for every pure module above (Vitest, next to the file). `npm run
build`, `npm run lint` on the new files, and a browser check of the page at phone
and desktop width. The page is behind the Microsoft sign-in, so the SharePoint
writes cannot be exercised end to end here; they are the existing, tested writers.

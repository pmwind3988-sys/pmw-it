# Till: guided "New item" sheet on stock-in

**Date:** 2026-11-17 · **Size:** small change · **Approved:** yes

## Why
Stocking in five identical laptops today means answering type, brand and model
five times: every new serial lands on the receipt as its own "New code" line.
The user wants the till to ask once — type, brand, model — then ask whether more
of the same product follow, each with its own serial, and take those serials as
the same product. And to add a new type from the till when the list lacks it.

## Approaches
- **Guided sheet opened by the scan (picked).** One question at a time, camera
  paused while typing so a stray read cannot land.
- Keep prompting inline on the receipt line — what exists now, and the cause of
  the repetition. Rejected.

## Design
1. Stock-in scan whose result is `NEW` (unknown code) opens a "New item" sheet.
2. **Type:** common-type chips, "More…" list, "Add a new type…" which writes the
   category to SharePoint immediately (`addCategory`, same as `CategoryField`).
3. **Brand:** free text, suggestions from brands in the register.
4. **Model:** free text, suggestions from register models of that type + brand.
   Picking a register model links the line (`linkToModel`).
5. **"More of the same product, each with its own serial?"**
   - Yes → `SerialRunSheet` opens; the first scanned serial is item 1.
     - Tracked type → each serial becomes its own draft of the same model.
     - Counted type → each serial is an item on the one line (`addSerialsTo`).
   - No → back to scanning.
6. **Skip** closes the sheet; the line keeps asking on the receipt as today.
   "Sweep the box" stays reachable from the sheet.

## Calls made (overridable)
- A type added at the till is counted (bulk), per `trackingModeFor`'s existing
  rule for unknown categories. Serials still live per item.

## Non-goals
- No per-type tracked/counted setting. No change to hand out / take back / count.
- Nothing reaches SharePoint except the new category; the receipt is still the
  IndexedDB batch.

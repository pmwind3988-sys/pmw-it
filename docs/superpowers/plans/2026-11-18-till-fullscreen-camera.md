# Till full-screen camera, tap to add — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** The till's camera fills the screen and lists what it reads in a card over the picture; only tapped codes go on the list.

**Architecture:** A pure pick list (`till/pickList.js`) decides which codes the card shows. `TillCamera` owns it and becomes a full-screen, body-portalled view; callers pass what to skip, how to describe a code, and what a tap does. The read gate and the "Which one did you mean?" chooser leave the camera path.

**Tech Stack:** React 19, Vitest, plain CSS (`src/styles/till.css`).

**Spec:** `docs/superpowers/specs/2026-11-18-till-fullscreen-camera-design.md`

## Global Constraints
- Shown only after 2 reads within 1500 ms; forgotten 4000 ms after last seen; at most 6 codes, oldest first.
- A tapped code reads "✓ added" and cannot be tapped again until forgotten.
- A code the mode would refuse as "already on this receipt" is never shown.
- No new dependencies; icons from `components/ui/Icons.jsx`.

## Review Focus
- A second identical box of mice (same shop code) must be addable once the first has left view for 4 s.
- A code seen once (misread) must never appear.
- The card must empty on its own when the camera points at nothing (no frames arrive then).
- New item / serial run sheets must open ABOVE the full-screen till camera; the serial run's own camera ABOVE its sheet.
- Escape / ✕ closes only the camera, never the sheet under it.

---

### Task 1: Pick list

**Files:** Create `src/features/assets/till/pickList.js`, `pickList.test.js`

**Interfaces — Produces:**
- `newPicks() -> Picks`
- `seePicks(picks, frameCodes, { now, skip = () => false }) -> Picks` — frameCodes are `{rawValue}` or strings, normalised with `normaliseCode`.
- `takePick(picks, code) -> Picks`
- `visiblePicks(picks, now) -> Array<{ code, taken }>`
- `PICK_TIMING = { need: 2, windowMs: 1500, keepMs: 4000, max: 6 }`

- [ ] Tests: `one read is not shown`; `two reads within 1.5s are shown`; `two reads 2s apart are not`; `a taken code reads taken and stays taken while in view`; `a code unseen for 4s disappears, and a taken one can be shown untaken after`; `skip hides a code`; `at most six, oldest first`.
- [ ] Run `npx vitest run src/features/assets/till/pickList.test.js` — FAIL, then implement, then PASS.

### Task 2: Full-screen TillCamera

**Files:** Modify `src/features/assets/till/ui/TillCamera.jsx`, `src/styles/till.css`

**Interfaces — Consumes:** Task 1. **Produces:** `<TillCamera active onTake(code) skip(code)->bool describe(code)->string flash onQuiet onClose status over />` — `status` is the corner text, `over` lifts it above till sheets.

- [ ] Portal to `document.body`; `.till-fullcam` fixed inset 0, z 65 (sheets are 70), `.till-fullcam-over` z 75; till colour tokens declared on it.
- [ ] Top bar: ✕ "Close camera" + status; flash under the bar; card at the bottom: "Seen on camera — tap to add", rows of code + describe, "✓ added" rows disabled; empty card says "Point at a barcode".
- [ ] 500 ms tick while any pick is visible; Escape (when active) calls onClose; body scroll locked while mounted.

### Task 3: Wire the till and the serial run

**Files:** Modify `src/pages/AssetTillPage.jsx`, `src/features/assets/till/ui/SerialRunSheet.jsx`

- [ ] Till: drop the gate, `choices`, `onCodes`; `skip` = dry run of `scanIn` / `scanOut` / `scanBack` (and stash machines) returning DUPLICATE; `onTake = handleCode`; `onClose` turns the camera off; status `"<n> on list"`.
- [ ] Serial run: camera on by default with `over`; `skip` = `addToRun` dry run returning REPEAT or ON_RECEIPT; `onClose` shows the sheet with a "Use the camera" button; status `"<n> scanned"`.
- [ ] `npx vitest run src/features/assets/till`, `npx eslint` on touched files, build; try in the app; update AGENTS.md's till paragraph; commit.

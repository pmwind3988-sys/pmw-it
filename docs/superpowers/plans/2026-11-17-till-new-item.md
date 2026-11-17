# Till "New item" sheet Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** An unknown stock-in scan opens a sheet asking type → brand → model → "more of the same, each with its own serial?", and a type can be added from it.

**Architecture:** Pure helpers in `till/stockIn.js` and `till/count.js` (tested with vitest); one new sheet `till/ui/NewItemSheet.jsx`; wiring in `pages/AssetTillPage.jsx`. The run reuses `SerialRunSheet`.

**Tech Stack:** React 19, Vite, vitest.

**Spec:** `docs/superpowers/specs/2026-11-17-till-new-item-design.md`

## Global Constraints
- Nothing but a new category reaches SharePoint; the receipt stays the IndexedDB batch.
- Inputs in the till stay ≥16px (iOS zoom rule in AGENTS.md).
- No helper exported beside a component (Fast Refresh rule).

## Review Focus
- First scanned code was a part number on a tracked type → the first run serial fills THAT line, not a new one.
- A run serial already on the receipt / in the register → refused by `addToRun` (existing).
- Skip → line still holds checkout with "Say what this is." as today.
- Picking a register model whose category matches → `linkToModel`, box remembered.
- A tracked run offers no "One without a serial" (a laptop with no serial goes through "Nothing on it").

---

### Task 1: Pure helpers
**Files:** Modify `src/features/assets/till/stockIn.js`, `src/features/assets/till/count.js`; Test `stockIn.test.js`, `count.test.js`

**Produces:**
- `nameLine(batch, localId, { category, manufacturer, model }) -> batch` — category via `setKind`, then make/model via `setDraftField` (blank values left alone).
- `addSameModel(batch, localId, run) -> batch`, `run = { serials, without, boxCodes }`. Tracked line: first serial fills the line when `needsSerial(line)`; every other serial becomes a new draft of the same category/make/model/part/additionalCodes/spec/location with `serialNumber`. Bulk line: `addSerialsTo`.
- `knownMakes(category, assets, drafts, limit = 8) -> string[]` — drafts newest first, then register by count; case-insensitive unique.

- [ ] Tests: nameLine makes a laptop line tracked with make/model and `needsKind` false; nameLine on Mouse moves the scanned serial onto item 1; addSameModel tracked with line serial `A1` and run `[A2,A3]` → 3 drafts same model, serials A1..A3; tracked line holding only a part number + `[S1,S2]` → line gets S1, one new draft S2; bulk → `serialCount` 3; knownMakes orders receipt first then register frequency.
- [ ] Run `npx vitest run src/features/assets/till` → FAIL, implement, → PASS. Commit.

### Task 2: NewItemSheet
**Files:** Create `src/features/assets/till/ui/NewItemSheet.jsx`; Modify `src/styles/till.css` if needed.
Props: `{ code, categories, assets, drafts, onName({category, manufacturer, model, asset}), onMore(), onSweep(), onSkip() }`. Steps: type (chips + More list + "Add a new type…" using `categoryRefusal` + `addCategory`), brand (input + `knownMakes` chips), model (input + `recentModels` filtered to brand; picking one whose register row exists passes `asset`), then the yes/no question. Camera paused while open.

### Task 3: Wire into the till
**Files:** Modify `src/pages/AssetTillPage.jsx`, `SerialRunSheet.jsx` (`allowWithout` prop, default true).
`STOCK_RESULT.NEW` sets `newFor = { localId, code }`; `onName` applies `linkToModel` (with asset) else `nameLine`; "Yes" opens `runFor` with `same: true`; run `onDone` uses `addSameModel` when `same`. Camera `active` also false while `newFor`. Verify in the browser at `/assets/till`; `npm run lint`, `npm run build`. Commit.

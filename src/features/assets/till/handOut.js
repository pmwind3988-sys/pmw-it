import { findScanTarget } from '../handover/scanMatch.js';
import {
  newLine, newUnitLine, addLine, hasAsset, hasUnit, setQuantity, isUnitLine,
} from '../handover/basket.js';
import { lineRefusal } from '../handover/planHandover.js';
import { HANDOVER_KIND } from '../handover/availability.js';
import { TRACKED } from '../assetKinds.js';
import { normaliseCode } from '../identity.js';

/**
 * Hand out at the till: what one scanned code does to the basket.
 *
 * The basket is the handover page's own (`handover/basket.js`) and the write
 * is its own (`commitHandover`); the till changes when a refusal is SEEN —
 * on the line, the moment the thing is scanned — not what counts as one.
 */

export const OUT_RESULT = {
  EMPTY: 'empty',
  ADDED: 'added',
  COUNTED: 'counted',
  DUPLICATE: 'duplicate',
  MISSING: 'missing',
};

/** A whole register row onto the basket: a search pick, a quick key, a scan. */
export function addAsset(basket, asset) {
  if (asset.trackingMode === TRACKED) {
    if (hasAsset(basket, asset.id)) return { basket, result: OUT_RESULT.DUPLICATE, line: null };
    const line = newLine(asset);
    return { basket: addLine(basket, line), result: OUT_RESULT.ADDED, line };
  }

  // A bulk row read again is one more of it, like a till.
  const existing = basket.lines.find((line) => line.assetId === asset.id && !isUnitLine(line));
  if (existing) {
    const next = setQuantity(basket, existing.lineId, existing.quantity + 1);
    return {
      basket: next,
      result: OUT_RESULT.COUNTED,
      line: next.lines.find((line) => line.lineId === existing.lineId),
    };
  }

  const line = newLine(asset, { quantity: 1 });
  return { basket: addLine(basket, line), result: OUT_RESULT.ADDED, line };
}

export function scanOut(basket, raw, assets = []) {
  if (!String(raw ?? '').trim()) return { basket, result: OUT_RESULT.EMPTY, line: null };

  const target = findScanTarget(assets, raw);
  if (!target) return { basket, result: OUT_RESULT.MISSING, line: null };

  const { asset, unit } = target;
  // Only a unit's OWN serial or label picks out one item. A part number is
  // printed on every box of the line, so reading it is "one more of these",
  // not "item 1" over and over.
  const code = normaliseCode(raw);
  const ownCode = unit
    && (normaliseCode(unit.serialNumber) === code || normaliseCode(unit.assetTag) === code);
  if (ownCode) {
    if (hasUnit(basket, asset.id, unit.index)) return { basket, result: OUT_RESULT.DUPLICATE, line: null };
    const line = newUnitLine(asset, unit);
    return { basket: addLine(basket, line), result: OUT_RESULT.ADDED, line };
  }
  return addAsset(basket, asset);
}

/** lineId → why it cannot go out, for every line that cannot. */
export function refusalsFor(basket, assets = []) {
  const byId = new Map(assets.map((asset) => [asset.id, asset]));
  const refusals = new Map();
  for (const line of basket.lines) {
    const reason = lineRefusal(line, byId.get(line.assetId), basket);
    if (reason) refusals.set(line.lineId, reason);
  }
  return refusals;
}

/** What is actually handed over: the lines that can go, nothing else. */
export function sendable(basket, refusals) {
  return { ...basket, lines: basket.lines.filter((line) => !refusals.has(line.lineId)) };
}

export function itemCount(basket, refusals = new Map()) {
  return basket.lines
    .filter((line) => !refusals.has(line.lineId))
    .reduce((sum, line) => sum + (line.quantity ?? 1), 0);
}

export const DUE_CHOICES = [
  { id: '1w', label: '1 week', days: 7 },
  { id: '2w', label: '2 weeks', days: 14 },
  { id: '1m', label: '1 month', months: 1 },
  { id: 'none', label: 'No date' },
];

/**
 * The due date for a choice, at LOCAL noon — never midnight UTC, which in
 * Malaysia is 8am the day before and stores as the wrong day.
 */
export function dueOnFor(choiceId, now = Date.now()) {
  const choice = DUE_CHOICES.find((entry) => entry.id === choiceId);
  if (!choice || (!choice.days && !choice.months)) return null;

  const date = new Date(now);
  date.setHours(12, 0, 0, 0);
  if (choice.days) date.setDate(date.getDate() + choice.days);
  if (choice.months) date.setMonth(date.getMonth() + choice.months);
  return date.getTime();
}

/** Loan or keep, and for how long. A kept thing has no date. */
export function withTerms(basket, { loan, dueChoice, now = Date.now() }) {
  return loan
    ? { ...basket, kind: HANDOVER_KIND.BORROWED, dueOn: dueOnFor(dueChoice, now) }
    : { ...basket, kind: HANDOVER_KIND.ISSUED, dueOn: null };
}

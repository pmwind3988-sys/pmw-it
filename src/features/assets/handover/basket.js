import { newId } from '../draft/draftAsset.js';
import { TRACKED } from '../assetKinds.js';
import { HANDOVER_KIND, available } from './availability.js';

/**
 * What is about to be handed to somebody.
 *
 * The person comes first and the lines are added to them, because a line cannot
 * be checked until it is known who it is for — and a refusal that only arrives
 * at the end is one nobody can act on while still standing at the desk (§4.4).
 */

export function newBasket(person = null) {
  return {
    id: newId(),
    person,
    // Set once for the whole basket and inherited by every line, the same
    // "undefined means inherit" contract the delivery batch uses.
    kind: HANDOVER_KIND.ISSUED,
    dueOn: null,
    remarks: '',
    lines: [],
  };
}

/**
 * One line. Tracked items are pinned to a quantity of one wherever they are
 * touched: there is one of them, and a handover of two would be a fiction.
 */
export function newLine(asset, overrides = {}) {
  return {
    lineId: newId(),
    assetId: asset.id,
    assetKey: asset.assetKey,
    itemTitle: asset.title ?? '',
    category: asset.category ?? '',
    trackingMode: asset.trackingMode,
    quantity: asset.trackingMode === TRACKED ? 1 : Math.min(1, available(asset)) || 1,
    kind: undefined,
    dueOn: undefined,
    remarks: '',
    ...overrides,
  };
}

const inherit = (lineValue, basketValue) => (lineValue === undefined ? basketValue : lineValue);

/** A line with the basket's kind and date filled in where it has none of its own. */
export function resolveLine(line, basket) {
  const kind = inherit(line.kind, basket.kind);
  return {
    ...line,
    kind,
    // An Issued item has no due date by definition, so one left over from a line
    // that used to be Borrowed is dropped rather than written.
    dueOn: kind === HANDOVER_KIND.BORROWED ? inherit(line.dueOn, basket.dueOn) : null,
  };
}

export function resolveLines(basket) {
  return (basket?.lines ?? []).map((line) => resolveLine(line, basket));
}

export function addLine(basket, line) {
  return { ...basket, lines: [...basket.lines, line] };
}

export function replaceLine(basket, line) {
  return {
    ...basket,
    lines: basket.lines.map((entry) => (entry.lineId === line.lineId ? line : entry)),
  };
}

export function removeLine(basket, lineId) {
  return { ...basket, lines: basket.lines.filter((line) => line.lineId !== lineId) };
}

export function setQuantity(basket, lineId, value) {
  const parsed = Number(value);
  const quantity = Number.isFinite(parsed) && parsed > 0 ? Math.floor(parsed) : 1;

  return {
    ...basket,
    lines: basket.lines.map((line) => (line.lineId === lineId
      ? { ...line, quantity: line.trackingMode === TRACKED ? 1 : quantity }
      : line)),
  };
}

/** Already in the basket, so the same laptop cannot be added to it twice. */
export function hasAsset(basket, assetId) {
  return basket.lines.some((line) => line.assetId === assetId);
}

export function unitCount(basket) {
  return basket.lines.reduce((sum, line) => sum + (line.quantity ?? 0), 0);
}

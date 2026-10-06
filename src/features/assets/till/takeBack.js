import { findScanTarget } from '../handover/scanMatch.js';
import { isOpen, outstanding, nameOfItem } from '../handover/availability.js';
import { newId } from '../draft/draftAsset.js';
import { normaliseCode } from '../identity.js';
import { unitsOf } from '../units.js';
import { TRACKED } from '../assetKinds.js';

/**
 * Take back at the till: a scanned thing becomes a return line for whoever
 * holds it, so nobody has to be looked up first.
 *
 * A line points at ONE open handover row, which is what `commitReturn` takes.
 * When several people hold the same bulk line the till cannot know whose came
 * back, so the line asks — it never guesses, because the wrong guess clears
 * somebody else's record.
 */

export const BACK_RESULT = {
  EMPTY: 'empty',
  ADDED: 'added',
  COUNTED: 'counted',
  DUPLICATE: 'duplicate',
  CHOOSE: 'choose',
  MISSING: 'missing',
  NOT_OUT: 'not-out',
  ALL_BACK: 'all-back',
};

function itemName(handover, asset) {
  const title = handover.itemTitle || asset?.title || 'Item';
  const which = nameOfItem(handover, asset ? unitsOf(asset) : []);
  return which ? `${title} · ${which}` : title;
}

/** A return line for one handover row. */
export function lineFor(handover, asset, quantity = 1) {
  const max = outstanding(handover);
  return {
    lineId: newId(),
    handoverId: handover.id,
    assetKey: handover.assetKey,
    name: itemName(handover, asset),
    from: handover.personName || handover.personEmail || 'Somebody',
    fromEmail: handover.personEmail || '',
    single: asset?.trackingMode === TRACKED || Number.isInteger(handover.unitIndex),
    quantity: Math.min(Math.max(1, quantity), max),
    max,
    choices: [],
  };
}

function choiceLine(open, asset) {
  return {
    lineId: newId(),
    handoverId: null,
    assetKey: asset.assetKey,
    name: asset.title || 'Item',
    from: '',
    fromEmail: '',
    single: false,
    quantity: 1,
    max: 1,
    choices: open.map((row) => ({ handoverId: row.id, from: row.personName || row.personEmail, outstanding: outstanding(row) })),
  };
}

/** The open handovers a scanned code can mean. */
export function openFor(raw, assets = [], handovers = []) {
  const target = findScanTarget(assets, raw);
  if (!target) return { target: null, open: [] };

  const { asset, unit } = target;
  let open = handovers.filter((row) => row.assetKey === asset.assetKey && isOpen(row));

  // A unit's own serial or label names one item, and only the handover of that
  // item can be it. A shared part number names the line.
  const code = normaliseCode(raw);
  if (unit && (normaliseCode(unit.serialNumber) === code || normaliseCode(unit.assetTag) === code)) {
    open = open.filter((row) => row.unitIndex === unit.index
      || (row.serialNumber && normaliseCode(row.serialNumber) === code));
  }
  return { target, open };
}

export function scanBack(lines, raw, assets = [], handovers = []) {
  if (!String(raw ?? '').trim()) return { lines, result: BACK_RESULT.EMPTY, line: null };

  const { target, open } = openFor(raw, assets, handovers);
  if (!target) return { lines, result: BACK_RESULT.MISSING, line: null };
  if (!open.length) return { lines, result: BACK_RESULT.NOT_OUT, line: null, asset: target.asset };

  // Already on the receipt for one of these: count up while there is more out.
  const existing = lines.find((line) => open.some((row) => row.id === line.handoverId));
  if (existing) {
    if (existing.single) return { lines, result: BACK_RESULT.DUPLICATE, line: existing };
    if (existing.quantity >= existing.max) return { lines, result: BACK_RESULT.ALL_BACK, line: existing };
    const next = { ...existing, quantity: existing.quantity + 1 };
    return { lines: lines.map((line) => (line.lineId === existing.lineId ? next : line)), result: BACK_RESULT.COUNTED, line: next };
  }

  if (open.length === 1) {
    const line = lineFor(open[0], target.asset, 1);
    return { lines: [...lines, line], result: BACK_RESULT.ADDED, line };
  }

  if (lines.some((line) => !line.handoverId && line.assetKey === target.asset.assetKey)) {
    return { lines, result: BACK_RESULT.DUPLICATE, line: null };
  }
  const line = choiceLine(open, target.asset);
  return { lines: [...lines, line], result: BACK_RESULT.CHOOSE, line };
}

/** A whole handover row onto the receipt, from search or "out right now". */
export function addHandover(lines, handover, assets = []) {
  if (lines.some((line) => line.handoverId === handover.id)) {
    return { lines, result: BACK_RESULT.DUPLICATE, line: null };
  }
  const asset = assets.find((entry) => entry.assetKey === handover.assetKey);
  const line = lineFor(handover, asset, outstanding(handover));
  return { lines: [...lines, line], result: BACK_RESULT.ADDED, line };
}

/** Answering "whose is it?" on a line several people could mean. */
export function chooseHolder(lines, lineId, handoverId, handovers = [], assets = []) {
  const row = handovers.find((entry) => entry.id === handoverId);
  if (!row) return lines;
  const asset = assets.find((entry) => entry.assetKey === row.assetKey);
  return lines.map((line) => (line.lineId === lineId ? { ...lineFor(row, asset, 1), lineId } : line));
}

export function setReturnQuantity(lines, lineId, value) {
  return lines.map((line) => {
    if (line.lineId !== lineId || line.single) return line;
    const quantity = Math.min(Math.max(1, Math.floor(Number(value) || 1)), line.max);
    return { ...line, quantity };
  });
}

/** What `commitReturn` takes. Lines still asking whose they are send nothing. */
export function toReturns(lines, condition) {
  return lines
    .filter((line) => line.handoverId != null)
    .map((line) => ({ handoverId: line.handoverId, quantity: line.quantity, condition }));
}

export function itemCount(lines) {
  return lines.filter((line) => line.handoverId != null).reduce((sum, line) => sum + line.quantity, 0);
}

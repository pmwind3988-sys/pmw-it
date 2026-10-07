import { normaliseCode } from '../identity.js';
import { TRACKED } from '../assetKinds.js';
import { unitsOf, parseUnits } from '../units.js';

/**
 * A serial run: ten of the same mouse, each with its own serial, scanned one
 * after another. Every serial is one item, so the count is simply how many
 * were scanned — nobody types a number. An item whose sticker is gone is
 * counted with `without` and still adds to the line.
 *
 * A serial can be on one thing only. One already in this run, on another line
 * of the receipt, or anywhere in the register is refused and named — a box
 * held in front of the camera a second time is the usual cause.
 */

export const RUN_RESULT = {
  ADDED: 'added',
  REPEAT: 'repeat',
  ON_RECEIPT: 'on-receipt',
  BOX_CODE: 'box-code',
  REGISTERED: 'registered',
  EMPTY: 'empty',
};

export function newRun() {
  return { serials: [], without: 0, boxCodes: [] };
}

/**
 * Not a serial: a shop barcode (12 to 14 digits, the same on every identical
 * box) or a barcode this line is already known by. It names the MODEL, so it
 * is kept once for the line and never becomes one item's serial.
 */
function isModelCode(code, known) {
  return /^\d{12,14}$/.test(code) || known.includes(code);
}

function inRegister(assets, code) {
  for (const asset of assets) {
    if (normaliseCode(asset.serialNumber) === code) return asset;
    if (asset.trackingMode !== TRACKED
      && unitsOf(asset).some((unit) => normaliseCode(unit.serialNumber) === code)) return asset;
  }
  return null;
}

function onReceipt(drafts, code) {
  return drafts.find((draft) => normaliseCode(draft.serialNumber) === code
    || parseUnits(draft.units).some((unit) => normaliseCode(unit.serialNumber) === code)) ?? null;
}

export function addToRun(run, raw, { assets = [], drafts = [], boxCodes = [] } = {}) {
  const code = normaliseCode(raw);
  if (!code) return { run, result: RUN_RESULT.EMPTY };
  const known = [...boxCodes, ...(run.boxCodes ?? [])].map(normaliseCode);
  if (isModelCode(code, known)) {
    const kept = known.includes(code);
    return { run: kept ? run : { ...run, boxCodes: [...(run.boxCodes ?? []), code] }, result: RUN_RESULT.BOX_CODE, code };
  }
  if (run.serials.includes(code)) return { run, result: RUN_RESULT.REPEAT };
  const draft = onReceipt(drafts, code);
  if (draft) return { run, result: RUN_RESULT.ON_RECEIPT, draft };
  const asset = inRegister(assets, code);
  if (asset) return { run, result: RUN_RESULT.REGISTERED, asset };
  return { run: { ...run, serials: [...run.serials, code] }, result: RUN_RESULT.ADDED };
}

export function withoutSerial(run) {
  return { ...run, without: run.without + 1 };
}

export function undoLast(run) {
  if (run.serials.length) return { ...run, serials: run.serials.slice(0, -1) };
  return { ...run, without: Math.max(0, run.without - 1) };
}

export function runSize(run) {
  return run.serials.length + run.without;
}

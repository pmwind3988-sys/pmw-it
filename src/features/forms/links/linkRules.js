import {
  ENTITIES, CHECKLIST_ITEMS, REQUESTABLE_ITEMS, fieldsFor,
} from '../checklistForm.js';
import { parseFormDate } from '../toChecklistItem.js';
import { LINK_STATUS } from './linkSchema.js';

/**
 * The rules of a shared checklist link, kept apart from SharePoint and React.
 *
 * Both sides read them: the public page to decide what to draw as an input,
 * and the server to decide what to believe. The server is the one that
 * matters — a browser can send anything, and "the employee cannot change the
 * laptop IT ticked" is only true if the server re-applies IT's value itself.
 */

/** Every field IT can pre-fill. The signature is never IT's to fill. */
export const LOCKABLE_FIELDS = [
  'employeeName', 'employeeNo', 'position', 'entity', 'formDate',
  'checkedItems', 'items', 'serialNumbers', 'otherRemarks',
];

// A submission still running after this long has died with its function, and
// the link goes back to waiting rather than staying stuck for good.
export const STALE_SIGNING_MS = 5 * 60 * 1000;

// About 1 MB of PNG once decoded. A drawn signature is tens of kilobytes; a
// body much larger than this is not a signature.
export const MAX_SIGNATURE_CHARS = 1_400_000;

const SHORT = 255;
const LONG = 4000;
const MAX_ITEM_LINES = 20;

const text = (value, limit) => String(value ?? '').trim().slice(0, limit);

export function isBlankValue(field, value) {
  if (field === 'items') {
    return !(value ?? []).some((row) => String(row?.item ?? '').trim());
  }
  if (Array.isArray(value)) return value.length === 0;
  return !String(value ?? '').trim();
}

/** A field IT left blank is always the employee's to fill. */
export function employeeMayEdit(link, field) {
  if (!LOCKABLE_FIELDS.includes(field)) return false;
  if (isBlankValue(field, link.preset?.[field])) return true;
  return (link.editable ?? []).includes(field);
}

/** The fields this link's form type shows that the employee may change. */
export function editableFields(link) {
  return fieldsFor(link.formMode)
    .filter((field) => LOCKABLE_FIELDS.includes(field))
    .filter((field) => employeeMayEdit(link, field));
}

export function linkState(link, now = Date.now()) {
  if (link.status === LINK_STATUS.SIGNED) return 'signed';
  if (link.status === LINK_STATUS.CANCELLED) return 'cancelled';

  const expires = Date.parse(link.expiresOn);
  if (!Number.isFinite(expires) || expires <= now) return 'expired';

  if (link.status === LINK_STATUS.SIGNING) {
    const since = Date.parse(link.modified);
    if (Number.isFinite(since) && now - since < STALE_SIGNING_MS) return 'busy';
  }
  return 'open';
}

function cleanItems(rows) {
  return (Array.isArray(rows) ? rows : [])
    .filter((row) => REQUESTABLE_ITEMS.includes(row?.item))
    .slice(0, MAX_ITEM_LINES)
    .map((row) => {
      const quantity = Math.floor(Number(row.quantity));
      return {
        item: row.item,
        quantity: Number.isFinite(quantity) ? Math.min(999, Math.max(1, quantity)) : 1,
      };
    });
}

function cleanSignature(value) {
  if (typeof value !== 'string' || value.length > MAX_SIGNATURE_CHARS) return null;
  return /^data:image\/png;base64,[A-Za-z0-9+/]+=*$/.test(value) ? value : null;
}

/**
 * What the browser sent, reduced to what the form could have produced: known
 * options only, bounded lengths, whole quantities. Anything else is dropped
 * rather than refused — validation then says what is missing in words.
 */
export function cleanSubmission(values = {}) {
  const day = text(values.formDate, 10);
  const checked = Array.isArray(values.checkedItems) ? values.checkedItems : [];

  return {
    employeeName: text(values.employeeName, SHORT),
    employeeNo: text(values.employeeNo, SHORT),
    position: text(values.position, SHORT),
    entity: ENTITIES.includes(values.entity) ? values.entity : '',
    formDate: parseFormDate(day) === null ? '' : day,
    checkedItems: CHECKLIST_ITEMS.filter((item) => checked.includes(item)),
    items: cleanItems(values.items),
    serialNumbers: text(values.serialNumbers, LONG),
    otherRemarks: text(values.otherRemarks, LONG),
    signature: cleanSignature(values.signature),
  };
}

/** IT's pre-filled values, cleaned the same way and without a signature. */
export function cleanPreset(values = {}) {
  const clean = cleanSubmission(values);
  delete clean.signature;
  return clean;
}

/**
 * The checklist as it will be signed: IT's value for every field the employee
 * may not change, the employee's for the rest, and the form type always from
 * the link.
 */
export function mergeSubmission(link, submitted) {
  const answer = cleanSubmission(submitted);
  const merged = { formMode: link.formMode, ...cleanPreset(link.preset) };

  for (const field of LOCKABLE_FIELDS) {
    if (employeeMayEdit(link, field)) merged[field] = answer[field];
  }
  merged.signature = answer.signature;
  return merged;
}

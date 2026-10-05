import { LINK_STATUS } from './linkSchema.js';
import { linkState, cleanPreset, editableFields } from './linkRules.js';
import { validateChecklist } from '../validate.js';
import { toChecklistItem } from '../toChecklistItem.js';
import { formatMYT } from '../../../utils/malaysiaTime.js';

/**
 * What IT can do to a link after sharing it — change when it expires, reopen
 * it, correct what was signed, remove it — as pure decisions. The portal
 * carries them out (`sharepoint/checklistLinks.js`); nothing here knows about
 * SharePoint or React.
 */

const ACTIONS = {
  open: ['copy', 'open', 'expiry', 'expireNow', 'cancel', 'delete'],
  // Somebody is submitting it this second. Changing it under them would make
  // their submission land on a link that no longer says what they saw.
  busy: ['open'],
  signed: ['view', 'edit', 'reopen', 'delete'],
  expired: ['reopen', 'delete'],
  cancelled: ['reopen', 'delete'],
};

export function linkActions(link, now = Date.now()) {
  return ACTIONS[linkState(link, now)] ?? [];
}

export function expiryFields(at) {
  const time = typeof at === 'number' ? at : Date.parse(at);
  if (!Number.isFinite(time)) throw new Error('That is not a date the link can expire on');
  return { ExpiresOn: new Date(time).toISOString() };
}

/**
 * A picked calendar day as the instant a link stops: the END of that day,
 * locally — "works until 19 October" means through the 19th, not up to its
 * first second.
 */
export function endOfDay(day) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(day ?? ''));
  if (!match) return NaN;
  const [, year, month, date] = match.map(Number);
  return new Date(year, month - 1, date, 23, 59, 59).getTime();
}

/** An instant as the `yyyy-mm-dd` a date input shows, in the local day. */
export function dayOf(at) {
  const time = typeof at === 'number' ? at : Date.parse(at);
  if (!Number.isFinite(time)) return '';
  const local = new Date(time);
  const pad = (value) => String(value).padStart(2, '0');
  return `${local.getFullYear()}-${pad(local.getMonth() + 1)}-${pad(local.getDate())}`;
}

/** A reopened link updates the record it already made instead of adding one. */
export const isReopened = (link) => Boolean(link.checklistId);

/**
 * Back to waiting, with a new expiry. A link that was signed reopens filled
 * with what was SIGNED — IT's corrections included — so what the employee
 * sees is the record as it stands, not IT's first draft of it. The signed row
 * and its signature are left exactly where they are until they sign again.
 */
export function reopenFields(link, expiresAt) {
  if (!link.submitted) return { LinkStatus: LINK_STATUS.WAITING, ...expiryFields(expiresAt) };

  // Everything the employee could change the first time stays theirs to
  // change. Without this, the answers they gave would come back as IT's
  // pre-filled values, locked, and a reopen could correct nothing.
  const editable = [...new Set([...(link.editable ?? []), ...editableFields(link)])];
  return {
    LinkStatus: LINK_STATUS.WAITING,
    ...expiryFields(expiresAt),
    Preset: JSON.stringify(link.submitted),
    Editable: JSON.stringify(editable),
  };
}

export function editNote(by, at) {
  return `Edited by ${by || 'IT'} on ${formatMYT(at, 'datetime12')} (Malaysia time), after signing`;
}

// The columns an edit may change. Not the signature, not the form type, and
// not when it was submitted: those are what the employee did.
const EDITABLE_COLUMNS = [
  'Title', 'EmployeeName', 'EmployeeNo', 'Position', 'Entity', 'Department',
  'FormDate', 'AssetMatrix', 'RequestedItems', 'SerialNumbers', 'OtherRemarks',
];

/**
 * IT correcting a checklist after the employee signed it.
 *
 * Allowed, and never silent: the row and the signed copy both say who edited
 * it and when, and the values as originally signed are kept on the link the
 * first time it is edited — later edits never overwrite that copy.
 */
export function planEdit(link, values, { by, at = Date.now() }) {
  if (link.status !== LINK_STATUS.SIGNED) {
    throw new Error('Only a signed checklist can be edited');
  }

  const clean = cleanPreset(values, link.options ?? null);
  const full = { ...clean, formMode: link.formMode };
  // The signature is not being changed, so it is not what is being checked.
  const errors = validateChecklist({ ...full, signature: 'kept' });

  const row = toChecklistItem(full);
  const checklist = Object.fromEntries(
    EDITABLE_COLUMNS.filter((column) => column in row).map((column) => [column, row[column]]),
  );
  checklist.EditedAfterSigning = editNote(by, at);

  const linkFields = {
    Submitted: JSON.stringify(clean),
    EmployeeName: clean.employeeName,
    EditedBy: by,
    EditedOn: new Date(at).toISOString(),
  };
  if (!link.originalSubmitted && link.submitted) {
    linkFields.OriginalSubmitted = JSON.stringify(link.submitted);
  }

  return { errors, values: clean, checklist, link: linkFields };
}

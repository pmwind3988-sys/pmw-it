import { isRequest } from './checklistForm.js';
import { formatMYT } from '../../utils/malaysiaTime.js';

/**
 * The checklist, as the row SharePoint stores.
 *
 * Items are written as readable lines — `Laptop x 1` — rather than as JSON.
 * This list is read by people in SharePoint, and a cell containing
 * `[{"item":"Laptop","quantity":1}]` is not a record anybody can use. The
 * portal is not the only consumer, so the storage has to be legible on its own.
 */

const text = (value) => String(value ?? '').trim();

export function formatItems(rows) {
  return (rows ?? [])
    .filter((row) => text(row.item))
    .map((row) => {
      const quantity = Number(row.quantity);
      const count = Number.isFinite(quantity) && quantity > 0 ? Math.floor(quantity) : 1;
      return `${text(row.item)} x ${count}`;
    })
    .join('\n');
}

export function formatChecked(items) {
  return (items ?? []).filter(Boolean).join('\n');
}

/**
 * `submittedAt` is the instant the form was actually sent; `formDate` is the
 * day the employee says the handover happened. They answer different questions
 * and are stored separately for that reason — on a signed record, "when was
 * this signed" and "when did he get the laptop" must not be confused.
 */
export function toChecklistItem(values, { submittedAt = Date.now(), signatureUrl = '' } = {}) {
  const item = {
    // Title is what the list shows as its link. Built from who and what, so a
    // row is identifiable without opening it.
    Title: [text(values.employeeName), text(values.employeeNo)].filter(Boolean).join(' — ')
      || 'Asset checklist',
    FormMode: values.formMode || '',
    EmployeeName: text(values.employeeName),
    EmployeeNo: text(values.employeeNo),
    Position: text(values.position),
    SerialNumbers: text(values.serialNumbers),
    OtherRemarks: text(values.otherRemarks),
    SignatureUrl: signatureUrl || '',
    SubmissionDate: new Date(submittedAt).toISOString(),
    SubmissionDateMYT: formatMYT(submittedAt, 'datetime12'),
    // Only one of the two is ever filled. The other is cleared rather than
    // left out, so re-reading a row never shows a previous shape's leftovers.
    AssetMatrix: isRequest(values.formMode) ? '' : formatChecked(values.checkedItems),
    RequestedItems: isRequest(values.formMode) ? formatItems(values.items) : '',
  };

  if (values.entity) item.Entity = values.entity;

  // A date-only input is a local day, and `new Date('2026-08-23')` reads it as
  // UTC midnight — which in Malaysia is the previous day at 8am. Parsed as
  // local noon instead, so the stored day is the day that was picked.
  const day = parseFormDate(values.formDate);
  if (day !== null) item.FormDate = new Date(day).toISOString();

  return item;
}

export function parseFormDate(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text(value));
  if (!match) return null;

  const [, year, month, dayOfMonth] = match;
  const parsed = new Date(Number(year), Number(month) - 1, Number(dayOfMonth), 12, 0, 0);
  return Number.isNaN(parsed.getTime()) ? null : parsed.getTime();
}

/** The signature file's name: identifiable, sortable, and legal on SharePoint. */
export function signatureFileName(values, submittedAt = Date.now()) {
  return `${values.formMode}-${new Date(submittedAt).toISOString()}-${values.entity}-${values.employeeName}.png`
    .replace(/[/\\?%*:|"<>]/g, '-')
    .replace(/\s+/g, '_');
}

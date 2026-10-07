import { FORM_MODES } from '../checklistForm.js';

/**
 * The list that holds shared checklist links, one row per link.
 *
 * The signed checklist itself is NOT kept here — it goes to `Asset Checklist
 * Form` exactly as one signed inside the portal does, so every view and report
 * on that list keeps working. This row is the link's own bookkeeping: what IT
 * pre-filled, what the employee may change, and where the signed row ended up.
 */

export const LINKS_LIST_NAME = 'Asset Checklist Links';

export const LINK_STATUS = {
  WAITING: 'Waiting',
  // Held for the few seconds a submission takes, so two taps — or two tabs —
  // cannot both sign the same link.
  SIGNING: 'Signing',
  SIGNED: 'Signed',
  CANCELLED: 'Cancelled',
};

const text = (StaticName, Title) => ({ StaticName, Title, kind: 'text' });
const note = (StaticName, Title) => ({ StaticName, Title, kind: 'note' });
const date = (StaticName, Title) => ({ StaticName, Title, kind: 'datetime' });
const number = (StaticName, Title) => ({ StaticName, Title, kind: 'number' });
const choice = (StaticName, Title, choices) => ({ StaticName, Title, kind: 'choice', choices });

export const LINK_COLUMNS = [
  choice('FormMode', 'Form Mode', FORM_MODES.map((mode) => mode.value)),
  text('EmployeeName', 'Employee Name'),
  note('Preset', 'Pre-filled values'),
  note('Editable', 'Employee may edit'),
  // HR's entities and departments as they stood when the link was made. The
  // public page has no way to read HR's lists itself (`formOptions.js`).
  note('Options', 'Dropdown choices'),
  date('ExpiresOn', 'Expires'),
  choice('LinkStatus', 'Status', Object.values(LINK_STATUS)),
  text('CreatedByName', 'Created by'),
  text('CreatedByEmail', 'Created by (email)'),
  note('Submitted', 'Signed values'),
  text('SignatureFile', 'Signature file'),
  number('ChecklistId', 'Checklist row'),
  date('SignedOn', 'Signed on'),
  // IT correcting what was signed. Never silent: who, when, and the values as
  // they were first signed, kept the first time and never overwritten.
  text('EditedBy', 'Edited by'),
  date('EditedOn', 'Edited on'),
  note('OriginalSubmitted', 'Values as first signed'),
  // The till's handover (or return) rows this checklist is the signature for:
  // { kind: 'issue' | 'return', ids: [handover row id] }. When the employee
  // signs, the server attaches that signature to those rows.
  note('Handovers', 'Till handover rows'),
];

export const LINK_VIEWS = [
  {
    list: LINKS_LIST_NAME,
    isDefault: true,
    title: 'All Items',
    fields: [
      'LinkTitle', 'EmployeeName', 'FormMode', 'LinkStatus', 'ExpiresOn',
      'CreatedByName', 'SignedOn', 'ChecklistId',
    ],
    query: '<OrderBy><FieldRef Name="Created" Ascending="FALSE" /></OrderBy>',
  },
];

const parse = (value, fallback, check) => {
  try {
    const parsed = JSON.parse(value);
    return check(parsed) ? parsed : fallback;
  } catch {
    return fallback;
  }
};

const isObject = (value) => value !== null && typeof value === 'object' && !Array.isArray(value);

/** A new link, as the row the portal creates. */
export function toLinkItem(link, { createdByName = '', createdByEmail = '' } = {}) {
  return {
    Title: link.code,
    FormMode: link.formMode,
    EmployeeName: String(link.preset?.employeeName ?? '').trim(),
    Preset: JSON.stringify(link.preset ?? {}),
    Editable: JSON.stringify(link.editable ?? []),
    Options: link.options ? JSON.stringify(link.options) : '',
    ExpiresOn: new Date(link.expiresOn).toISOString(),
    LinkStatus: link.status ?? LINK_STATUS.WAITING,
    CreatedByName: createdByName,
    CreatedByEmail: createdByEmail,
    Handovers: link.handovers ? JSON.stringify(link.handovers) : '',
  };
}

/**
 * A row back to a link. Reads the fields as either SharePoint REST or Graph
 * returns them — the portal reads over one and the server over the other, and
 * both use the same internal names.
 */
export function fromLinkItem(fields = {}) {
  const id = Number(fields.id ?? fields.Id ?? fields.ID);
  const checklistId = Number(fields.ChecklistId);

  return {
    id: Number.isFinite(id) ? id : null,
    code: fields.Title ?? '',
    formMode: fields.FormMode ?? '',
    employeeName: fields.EmployeeName ?? '',
    preset: parse(fields.Preset, {}, isObject),
    editable: parse(fields.Editable, [], Array.isArray),
    options: parse(fields.Options, null, isObject),
    expiresOn: fields.ExpiresOn ?? '',
    status: fields.LinkStatus || LINK_STATUS.WAITING,
    createdByName: fields.CreatedByName ?? '',
    createdByEmail: fields.CreatedByEmail ?? '',
    created: fields.Created ?? '',
    modified: fields.Modified ?? '',
    submitted: parse(fields.Submitted, null, isObject),
    signatureFile: fields.SignatureFile ?? '',
    checklistId: Number.isFinite(checklistId) && checklistId > 0 ? checklistId : null,
    signedOn: fields.SignedOn ?? '',
    editedBy: fields.EditedBy ?? '',
    editedOn: fields.EditedOn ?? '',
    originalSubmitted: parse(fields.OriginalSubmitted, null, isObject),
    handovers: parse(fields.Handovers, null, isObject),
  };
}

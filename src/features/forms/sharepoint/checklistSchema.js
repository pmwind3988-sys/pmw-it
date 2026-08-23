import { ENTITIES, FORM_MODES } from '../checklistForm.js';

export const CHECKLIST_LIST_NAME = 'Asset Checklist Form';
export const SIGNATURE_LIBRARY_NAME = 'Signatures';

const text = (StaticName, Title) => ({ StaticName, Title, kind: 'text' });
const note = (StaticName, Title) => ({ StaticName, Title, kind: 'note' });
const date = (StaticName, Title) => ({ StaticName, Title, kind: 'datetime' });
const choice = (StaticName, Title, choices) => ({ StaticName, Title, kind: 'choice', choices });

/**
 * The signed checklist, as SharePoint holds it.
 *
 * The first eight columns are what the list already had; the last four are new.
 * Nothing is renamed or removed, so every checklist already signed reads
 * exactly as it did.
 *
 * Provisioning goes through the shared `provisionSchema`, which replaces
 * `ensureAssetColumns` in `sharePointService.js` — that copy sent `Choices` on
 * a base `SP.Field` and would fail outright on a fresh site.
 */
export const CHECKLIST_COLUMNS = [
  choice('FormMode', 'Form Mode', FORM_MODES.map((mode) => mode.value)),
  text('EmployeeName', 'Employee Name'),
  text('EmployeeNo', 'Employee No'),
  text('Position', 'Position'),
  // Only ever ADDED to. The options this list already carries stay, or every
  // row saved with one of them becomes unreadable in its own list.
  choice('Entity', 'Entity', ENTITIES),
  date('SubmissionDate', 'Submission Date'),
  note('AssetMatrix', 'Asset Checklist'),
  text('SignatureUrl', 'Signature URL'),

  // The day the employee says the handover happened, which is theirs to set —
  // as opposed to `SubmissionDate`, which is when the form was actually sent.
  date('FormDate', 'Date'),
  note('RequestedItems', 'Requested Items'),
  note('SerialNumbers', 'Serial Numbers'),
  note('OtherRemarks', 'Other Remarks'),
  text('SubmissionDateMYT', 'Submitted (MYT)'),
];

const NAME = 'LinkTitle';

/**
 * REST-created columns join no view, so without this a freshly provisioned list
 * shows nothing but its Title.
 */
export const CHECKLIST_VIEWS = [
  {
    list: CHECKLIST_LIST_NAME,
    isDefault: true,
    title: 'All Items',
    fields: [
      NAME, 'FormMode', 'EmployeeName', 'EmployeeNo', 'Position', 'Entity',
      'FormDate', 'AssetMatrix', 'RequestedItems', 'SerialNumbers',
      'OtherRemarks', 'SignatureUrl', 'SubmissionDateMYT',
    ],
    query: '<OrderBy><FieldRef Name="SubmissionDate" Ascending="FALSE" /></OrderBy>',
  },
];

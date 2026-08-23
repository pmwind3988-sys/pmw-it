import {
  ASSET_LIST_NAME, BATCH_LIST_NAME, CHANGE_LIST_NAME, ASSET_COLUMNS,
} from './assetSchema.js';
import { HANDOVER_LIST_NAME } from './handoverSchema.js';

/**
 * `LinkTitle` rather than `Title`: it renders the name as the link into the
 * item, which is what makes a SharePoint view navigable.
 */
const NAME = 'LinkTitle';

/**
 * Bookkeeping that is on every row and tells nobody anything at a glance. The
 * item page still shows them; the wall of columns does not need to.
 */
const HIDDEN_FROM_ALL_ITEMS = new Set(['GuessedFields', 'ManualFields', 'AssetKey']);

const EVERY_ASSET_FIELD = [
  NAME,
  ...ASSET_COLUMNS
    .map((column) => column.StaticName)
    .filter((name) => !HIDDEN_FROM_ALL_ITEMS.has(name)),
];

/**
 * REST-created columns join no view, so a freshly provisioned list shows
 * nothing but its Title until these are set. These are the cuts the register
 * is actually read for.
 */
export const ASSET_VIEWS = [
  {
    list: ASSET_LIST_NAME,
    isDefault: true,
    title: 'All Items',
    fields: EVERY_ASSET_FIELD,
  },
  {
    list: ASSET_LIST_NAME,
    title: 'In stock',
    fields: [
      NAME, 'Category', 'Manufacturer', 'Model', 'Quantity', 'Condition',
      'Location', 'AssetTag', 'ArrivedOn',
    ],
    query:
      '<Where><Eq><FieldRef Name="Status" /><Value Type="Text">In stock</Value></Eq></Where>'
      + '<OrderBy><FieldRef Name="ArrivedOn" Ascending="FALSE" /></OrderBy>',
  },
  {
    // "What still needs a sticker" — the question §4.7 promises is one click.
    // Tracked rows only: a bag of cables was never going to carry a label.
    list: ASSET_LIST_NAME,
    title: 'Needs a label',
    fields: [NAME, 'Category', 'Manufacturer', 'Model', 'SerialNumber', 'Location', 'ArrivedOn'],
    query:
      '<Where><And>'
      + '<Eq><FieldRef Name="TrackingMode" /><Value Type="Text">Tracked</Value></Eq>'
      + '<IsNull><FieldRef Name="AssetTag" /></IsNull>'
      + '</And></Where>'
      + '<OrderBy><FieldRef Name="ArrivedOn" Ascending="FALSE" /></OrderBy>',
  },
  {
    list: ASSET_LIST_NAME,
    title: 'Recent arrivals',
    fields: [
      NAME, 'Category', 'Manufacturer', 'Model', 'Quantity', 'Supplier',
      'PoNumber', 'ArrivedOnMYT', 'BatchTitle',
    ],
    query: '<OrderBy><FieldRef Name="ArrivedOn" Ascending="FALSE" /></OrderBy>',
  },
  {
    list: ASSET_LIST_NAME,
    title: 'Out on loan',
    fields: [
      NAME, 'Category', 'AssignedTo', 'HandoverKind', 'DueOn', 'Quantity',
      'QuantityOut', 'SerialNumber', 'AssetTag',
    ],
    query:
      '<Where><Or>'
      + '<Eq><FieldRef Name="Status" /><Value Type="Text">Assigned</Value></Eq>'
      + '<Eq><FieldRef Name="Status" /><Value Type="Text">Borrowed</Value></Eq>'
      + '</Or></Where>'
      + '<OrderBy><FieldRef Name="DueOn" Ascending="TRUE" /></OrderBy>',
  },
  {
    list: HANDOVER_LIST_NAME,
    isDefault: true,
    title: 'All Items',
    fields: [
      NAME, 'ItemTitle', 'Category', 'PersonName', 'PersonEmail', 'Quantity',
      'ReturnedQuantity', 'Kind', 'HandoverStatus', 'IssuedOnMYT', 'DueOnMYT',
      'ReturnedOnMYT', 'ReturnCondition', 'IssuedBy', 'Remarks',
    ],
    query: '<OrderBy><FieldRef Name="IssuedOn" Ascending="FALSE" /></OrderBy>',
  },
  {
    list: HANDOVER_LIST_NAME,
    title: 'Currently out',
    fields: [
      NAME, 'ItemTitle', 'PersonName', 'Quantity', 'ReturnedQuantity', 'Kind',
      'DueOnMYT', 'IssuedOnMYT',
    ],
    query:
      '<Where><Neq><FieldRef Name="HandoverStatus" /><Value Type="Text">Returned</Value></Neq></Where>'
      + '<OrderBy><FieldRef Name="IssuedOn" Ascending="FALSE" /></OrderBy>',
  },
  {
    // The overdue view lives on the HANDOVER list rather than the register: a
    // box of cables held by three people has no single due date of its own,
    // so the answer only exists per handover.
    list: HANDOVER_LIST_NAME,
    title: 'Overdue',
    fields: [NAME, 'ItemTitle', 'PersonName', 'PersonEmail', 'Quantity', 'DueOnMYT'],
    query:
      '<Where><And>'
      + '<Neq><FieldRef Name="HandoverStatus" /><Value Type="Text">Returned</Value></Neq>'
      + '<Lt><FieldRef Name="DueOn" /><Value Type="DateTime"><Today /></Value></Lt>'
      + '</And></Where>'
      + '<OrderBy><FieldRef Name="DueOn" Ascending="TRUE" /></OrderBy>',
  },
  {
    list: BATCH_LIST_NAME,
    isDefault: true,
    title: 'All Items',
    fields: [
      NAME, 'Supplier', 'PoNumber', 'ItemCount', 'ArrivedOnMYT', 'PoPhotoUrl',
      'SavedOn', 'SavedBy', 'Remarks',
    ],
    query: '<OrderBy><FieldRef Name="ArrivedOn" Ascending="FALSE" /></OrderBy>',
  },
  {
    list: CHANGE_LIST_NAME,
    isDefault: true,
    title: 'All Items',
    fields: [NAME, 'FieldName', 'OldValue', 'NewValue', 'ChangeType', 'ChangedOn', 'ChangedBy'],
    query: '<OrderBy><FieldRef Name="ChangedOn" Ascending="FALSE" /></OrderBy>',
  },
];

import { formatMYT } from '../../datastudio/time/malaysiaTime.js';
import { CATEGORIES, CONDITIONS } from '../assetKinds.js';
import { HANDOVER_KIND, HANDOVER_STATUS } from '../handover/availability.js';

export const HANDOVER_LIST_NAME = 'IT Asset Handovers';

const text = (StaticName, Title) => ({ StaticName, Title, kind: 'text' });
const note = (StaticName, Title) => ({ StaticName, Title, kind: 'note' });
const num = (StaticName, Title) => ({ StaticName, Title, kind: 'number' });
const date = (StaticName, Title) => ({ StaticName, Title, kind: 'datetime' });
const choice = (StaticName, Title, choices) => ({ StaticName, Title, kind: 'choice', choices });

/**
 * One row per item handed to somebody. THIS list is the truth about who has
 * what; the copies on the register row exist only so that a row opened
 * directly in SharePoint reads without a join (§4.2).
 *
 * `Title` is "<person> — <item>", readable and never load-bearing.
 */
export const HANDOVER_COLUMNS = [
  text('HandoverId', 'Handover'),
  text('AssetKey', 'Asset Key'),
  num('AssetId', 'Asset Row'),
  text('ItemTitle', 'Item'),
  choice('Category', 'Category', CATEGORIES),

  text('PersonName', 'Person'),
  // The identity. "What does Amir have" keys on this, never on the display
  // name, which two people will spell two ways.
  text('PersonEmail', 'Person Email'),
  text('PersonLogin', 'Person Login'),

  num('Quantity', 'Quantity'),
  num('ReturnedQuantity', 'Returned'),
  choice('Kind', 'Kind', [HANDOVER_KIND.ISSUED, HANDOVER_KIND.BORROWED]),
  choice('HandoverStatus', 'Status', [
    HANDOVER_STATUS.OUT, HANDOVER_STATUS.PARTLY, HANDOVER_STATUS.RETURNED,
  ]),

  date('IssuedOn', 'Issued On'),
  text('IssuedOnMYT', 'Issued On (MYT)'),
  date('DueOn', 'Due Back'),
  text('DueOnMYT', 'Due Back (MYT)'),
  date('ReturnedOn', 'Returned On'),
  text('ReturnedOnMYT', 'Returned On (MYT)'),
  choice('ReturnCondition', 'Condition Returned', CONDITIONS),

  text('IssuedBy', 'Issued By'),
  text('ReturnedBy', 'Returned By'),
  note('Remarks', 'Remarks'),
];

const keyFor = (staticName) => staticName.charAt(0).toLowerCase() + staticName.slice(1);

/**
 * Guarded, because `formatMYT` throws on a value that is not an instant — and
 * an issued item legitimately has no due date and no return date. Unguarded,
 * every ordinary handover would fail to write.
 */
function readableMYT(value) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '';
  return formatMYT(value, 'datetime12');
}

export function toListItem(handover) {
  const item = { Title: handover.title ?? '' };

  for (const column of HANDOVER_COLUMNS) {
    const key = keyFor(column.StaticName);
    let value = handover[key];

    if (column.StaticName === 'IssuedOnMYT') value = readableMYT(handover.issuedOn);
    if (column.StaticName === 'DueOnMYT') value = readableMYT(handover.dueOn);
    if (column.StaticName === 'ReturnedOnMYT') value = readableMYT(handover.returnedOn);

    switch (column.kind) {
      case 'text':
      case 'note':
        item[column.StaticName] = value == null ? '' : String(value);
        break;
      case 'number':
        // `0` is a real returned-quantity and must survive.
        if (typeof value === 'number' && Number.isFinite(value)) item[column.StaticName] = value;
        break;
      case 'choice':
        if (value) item[column.StaticName] = String(value);
        break;
      case 'datetime':
        if (typeof value === 'number' && Number.isFinite(value)) {
          item[column.StaticName] = new Date(value).toISOString();
        }
        break;
      default:
        break;
    }
  }

  return item;
}

export function fromListItem(row) {
  const record = { id: row.Id ?? row.ID ?? null, title: row.Title ?? null };

  for (const column of HANDOVER_COLUMNS) {
    const key = keyFor(column.StaticName);
    const raw = row[column.StaticName];

    if (raw === undefined || raw === null || raw === '') {
      record[key] = null;
      continue;
    }

    record[key] = column.kind === 'datetime' ? new Date(raw).getTime() : raw;
  }

  // Both are summed and compared everywhere; a null would poison the
  // arithmetic in a way that reads as "nothing is out" rather than as an error.
  if (record.quantity == null) record.quantity = 1;
  if (record.returnedQuantity == null) record.returnedQuantity = 0;

  return record;
}

/**
 * A MERGE body for a return. Built from the same column definitions rather than
 * hand-typed, so a renamed column cannot leave this silently writing nothing.
 */
export function toUpdateItem(patch) {
  const item = {};

  for (const column of HANDOVER_COLUMNS) {
    const key = keyFor(column.StaticName);
    if (!(key in patch)) continue;

    const value = patch[key];

    if (column.kind === 'datetime') {
      if (typeof value === 'number' && Number.isFinite(value)) {
        item[column.StaticName] = new Date(value).toISOString();
      } else {
        item[column.StaticName] = null;
      }
      continue;
    }

    if (column.kind === 'number') {
      if (typeof value === 'number' && Number.isFinite(value)) item[column.StaticName] = value;
      continue;
    }

    if (column.kind === 'choice') {
      if (value) item[column.StaticName] = String(value);
      continue;
    }

    item[column.StaticName] = value == null ? '' : String(value);
  }

  if ('returnedOn' in patch) item.ReturnedOnMYT = readableMYT(patch.returnedOn);

  return item;
}

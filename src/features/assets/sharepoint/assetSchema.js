import { formatMYT } from '../../datastudio/time/malaysiaTime.js';
import { CATEGORIES, CONDITIONS, STATUSES, TRACKED, BULK } from '../assetKinds.js';

export const ASSET_LIST_NAME = 'IT Asset Register';
export const BATCH_LIST_NAME = 'IT Asset Batches';
export const CHANGE_LIST_NAME = 'IT Asset Changes';
export const PHOTO_LIBRARY_NAME = 'IT Asset Photos';

const text = (StaticName, Title) => ({ StaticName, Title, kind: 'text' });
const note = (StaticName, Title) => ({ StaticName, Title, kind: 'note' });
const num = (StaticName, Title) => ({ StaticName, Title, kind: 'number' });
const date = (StaticName, Title) => ({ StaticName, Title, kind: 'datetime' });
const choice = (StaticName, Title, choices) => ({ StaticName, Title, kind: 'choice', choices });

/**
 * `Title` is built in and holds a readable name. Identity lives in `AssetKey`
 * — a name that changes when somebody corrects a model number must not move
 * the row.
 */
export const ASSET_COLUMNS = [
  text('AssetKey', 'Asset Key'),
  choice('Category', 'Category', CATEGORIES),
  choice('TrackingMode', 'Tracking', [TRACKED, BULK]),

  text('Manufacturer', 'Manufacturer'),
  text('Model', 'Model'),
  text('SerialNumber', 'Serial Number'),
  text('PartNumber', 'Part Number'),
  text('MacAddress', 'MAC Address'),
  note('AdditionalCodes', 'Other Codes'),

  text('AssetTag', 'Asset Label'),
  num('Quantity', 'Quantity'),
  choice('Condition', 'Condition', CONDITIONS),
  choice('Status', 'Status', STATUSES),
  text('Location', 'Location'),
  note('Remarks', 'Remarks'),
  note('SpecSummary', 'Specification'),

  text('Supplier', 'Purchased From'),
  text('PoNumber', 'PO Number'),
  date('ArrivedOn', 'Arrived On'),
  text('ArrivedOnMYT', 'Arrived On (MYT)'),
  date('PurchasedOn', 'Purchased On'),

  text('BatchId', 'Delivery ID'),
  text('BatchTitle', 'Delivery'),
  text('PhotoUrl', 'Photo'),
  text('PoPhotoUrl', 'PO Scan'),

  choice('ScanSource', 'Added By', ['Camera', 'Manual']),
  note('GuessedFields', 'Guessed Fields'),
  note('ManualFields', 'Manually Set Fields'),

  // Provisioned now, filled by the assignment project. Adding a column to a
  // list with rows in it is a migration; adding it up front is a line here.
  text('AssignedTo', 'Assigned To'),
  text('AssignedToEmail', 'Assigned To (email)'),
  date('AssignedOn', 'Assigned On'),

  date('AddedOn', 'Added On'),
  text('AddedOnMYT', 'Added On (MYT)'),
  text('AddedBy', 'Added By (user)'),
];

export const BATCH_COLUMNS = [
  text('Supplier', 'Purchased From'),
  text('PoNumber', 'PO Number'),
  date('ArrivedOn', 'Arrived On'),
  text('ArrivedOnMYT', 'Arrived On (MYT)'),
  text('PoPhotoUrl', 'PO Scan'),
  num('ItemCount', 'Items'),
  note('Remarks', 'Remarks'),
  date('SavedOn', 'Saved On'),
  text('SavedBy', 'Saved By'),
];

export const CHANGE_COLUMNS = [
  text('FieldName', 'Field'),
  note('OldValue', 'Old Value'),
  note('NewValue', 'New Value'),
  date('ChangedOn', 'Changed On'),
  text('ChangedOnMYT', 'Changed On (MYT)'),
  text('ChangedBy', 'Changed By'),
  choice('ChangeType', 'Change Type', ['Added', 'Updated', 'Removed']),
];

/**
 * Only these produce change-log rows. `AdditionalCodes` and the photo URLs are
 * deliberately absent: they churn every time something is re-scanned or
 * re-photographed, and logging them would bury the changes that matter — a
 * label moving to another machine, a quantity dropping, a status changing.
 */
export const TRACKED_FIELDS = [
  'category', 'trackingMode', 'manufacturer', 'model',
  'serialNumber', 'partNumber', 'assetTag',
  'quantity', 'condition', 'status', 'location',
  'supplier', 'poNumber', 'assignedTo',
];

/** camelCase record key for a StaticName: first letter lowered. */
const keyFor = (staticName) => staticName.charAt(0).toLowerCase() + staticName.slice(1);

const serialise = (value) => {
  if (Array.isArray(value)) return value.map((entry) => String(entry)).join('\n');
  return value == null ? '' : String(value);
};

function readableMYT(value) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '';
  return formatMYT(value, 'datetime12');
}

export function toListItem(asset) {
  const item = { Title: asset.title ?? '' };

  for (const column of ASSET_COLUMNS) {
    const key = keyFor(column.StaticName);
    let value = asset[key];

    // Guarded, because `formatMYT` throws on a value that is not an instant
    // and an item can perfectly well arrive without anybody recording when.
    // Unguarded, one undated row takes the whole save down with it.
    if (column.StaticName === 'ArrivedOnMYT') value = readableMYT(asset.arrivedOn);
    if (column.StaticName === 'AddedOnMYT') value = readableMYT(asset.addedOn);

    switch (column.kind) {
      case 'text':
      case 'note':
        // Empty string clears the column; null would be rejected.
        item[column.StaticName] = serialise(value);
        break;
      case 'number':
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

const ARRAY_COLUMNS = new Set(['AdditionalCodes', 'GuessedFields', 'ManualFields']);

export function fromListItem(row) {
  const record = { id: row.Id ?? row.ID ?? null, title: row.Title ?? null };

  for (const column of ASSET_COLUMNS) {
    const key = keyFor(column.StaticName);
    const raw = row[column.StaticName];

    // An absent column reads as null for every kind — notably NOT as NaN for a
    // date, which is what `new Date(undefined).getTime()` would produce.
    if (raw === undefined || raw === null || raw === '') {
      record[key] = ARRAY_COLUMNS.has(column.StaticName) ? [] : null;
      continue;
    }

    if (column.kind === 'datetime') record[key] = new Date(raw).getTime();
    else if (ARRAY_COLUMNS.has(column.StaticName)) record[key] = String(raw).split('\n');
    else record[key] = raw;
  }

  // Quantity is what every count on the page sums. A row saved before the
  // column existed reads as null, and `null + 1` is 1 — so it defaults here,
  // once, rather than at each of the places that add it up.
  if (record.quantity == null) record.quantity = 1;

  return record;
}

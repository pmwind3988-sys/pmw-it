import { spFetch, getFormDigest, listPath } from './spClient.js';
import {
  DEVICE_COLUMNS, CHANGE_COLUMNS, DEVICE_LIST_NAME, CHANGE_LIST_NAME,
} from './deviceSchema.js';

const FIELD_TYPE_KIND = { text: 2, note: 3, datetime: 4, choice: 6, boolean: 8, number: 9 };

/**
 * Each column is created as its own concrete type, not the base `SP.Field`.
 * Verified against the tenant: `SP.Field` does not declare `Choices`, and
 * sending it there fails with
 *   "The property 'Choices' does not exist on type 'SP.Field'".
 * The same reasoning applies to DisplayFormat, RichText and the rest — a
 * property only exists on the type that declares it.
 */
const METADATA_TYPE = {
  text: 'SP.Field',
  boolean: 'SP.Field',
  choice: 'SP.FieldChoice',
  note: 'SP.FieldMultiLineText',
  datetime: 'SP.FieldDateTime',
  number: 'SP.FieldNumber',
};

/**
 * Created under the internal name, NOT the display name.
 *
 * SharePoint derives a field's internal name from the `Title` it is created
 * with — `StaticName` in the creation body does not control it. So a field
 * created as "Device Type" is addressable only as `Device_x0020_Type`, and
 * writing `DeviceType` on the item fails with "The property 'DeviceType' does
 * not exist on type 'SP.Data...ListItem'".
 *
 * Creating it as `DeviceType` and renaming it afterwards (see `renameBody`)
 * gives a clean internal name AND a readable column header. This is the same
 * trap that produced the hand-encoded `Calling_x0020_Name` in
 * `src/services/sharePointService.js`.
 */
export function fieldBody(column) {
  const body = {
    __metadata: { type: METADATA_TYPE[column.kind] },
    Title: column.StaticName,
    StaticName: column.StaticName,
    FieldTypeKind: FIELD_TYPE_KIND[column.kind],
    Required: false,
  };

  // DisplayFormat means different things per field type: on DateTime, 1 keeps
  // the time (0 would be date-only and would discard it); on Number, 0 means
  // zero decimal places.
  if (column.kind === 'datetime') body.DisplayFormat = 1;
  if (column.kind === 'number') body.DisplayFormat = 0;

  if (column.kind === 'note') {
    // A rich-text Note stores <div> markup around the value, so a multi-answer
    // field would not round-trip.
    body.RichText = false;
    body.AppendOnly = false;
    body.NumberOfLines = 6;
  }

  if (column.kind === 'choice') body.Choices = { results: column.choices };

  return body;
}

/** The second half of the create-then-rename dance: set the display name. */
export function renameBody(column) {
  return { __metadata: { type: METADATA_TYPE[column.kind] }, Title: column.Title };
}

async function ensureList(siteUrl, token, digest, title, description) {
  const existing = await spFetch(siteUrl, listPath(title), { token });
  if (existing.ok) return;
  if (existing.status !== 404) {
    throw new Error(`Could not check for the "${title}" list (${existing.status})`);
  }

  const created = await spFetch(siteUrl, '/_api/web/lists', {
    token,
    digest,
    method: 'POST',
    body: {
      __metadata: { type: 'SP.List' },
      BaseTemplate: 100,
      Title: title,
      Description: description,
    },
  });

  if (!created.ok && created.status !== 201) {
    throw new Error(
      `Could not create the "${title}" list (${created.status}): ${await created.text()}`,
    );
  }
}

/**
 * Keyed by internal name, because that is the name item writes address.
 * `StaticName` is the wrong key here: it can disagree with the internal name,
 * and a column where the two disagree is exactly the broken case below.
 */
async function existingFields(siteUrl, token, title) {
  const path = `${listPath(title)}/fields?$select=InternalName,Title`;
  const response = await spFetch(siteUrl, path, { token });
  if (!response.ok) throw new Error(`Could not read the fields of "${title}" (${response.status})`);

  const data = await response.json();
  return new Map((data.d?.results ?? []).map((field) => [field.InternalName, field.Title]));
}

/**
 * A leftover from before columns were created under their internal names: the
 * header we want, sitting on a field the item writes cannot reach.
 *
 * Its internal name is always an encoded one, since encoding is the only thing
 * that makes the two names diverge. Requiring that keeps a built-in field which
 * happens to share a display name from being mistaken for one of these.
 */
function staleColumn(column, fields) {
  for (const [internalName, display] of fields) {
    if (display === column.Title && internalName.includes('_x00')) return internalName;
  }
  return null;
}

/** Display name only. The internal name is fixed at creation and stays put. */
async function renameColumn(siteUrl, token, digest, title, column) {
  const renamed = await spFetch(
    siteUrl,
    `${listPath(title)}/fields/getByInternalNameOrTitle('${column.StaticName}')`,
    {
      token,
      digest,
      method: 'POST',
      body: renameBody(column),
      headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
    },
  );

  if (!renamed.ok) {
    throw new Error(
      `Could not rename the "${column.StaticName}" column (${renamed.status}): `
        + `${await renamed.text()}`,
    );
  }
}

/** One column, from whatever state it is currently in to correct. */
async function ensureColumn(siteUrl, token, digest, title, column, fields) {
  if (fields.has(column.StaticName)) {
    // Finishes a run that died between creating a column and renaming it,
    // which would otherwise leave the header reading `DeviceType` for good.
    if (fields.get(column.StaticName) !== column.Title) {
      await renameColumn(siteUrl, token, digest, title, column);
    }
    return;
  }

  const stale = staleColumn(column, fields);
  if (stale) {
    throw new Error(
      `The "${title}" list shows "${column.Title}" on a column named ${stale}, which `
        + 'item writes cannot address. Delete that column and save again.',
    );
  }

  const response = await spFetch(siteUrl, `${listPath(title)}/fields`, {
    token,
    digest,
    method: 'POST',
    body: fieldBody(column),
  });

  // 409 means another tab won the race. The column exists either way.
  if (!response.ok && response.status !== 409) {
    throw new Error(
      `Could not create the "${column.Title}" column (${response.status}): `
        + `${await response.text()}`,
    );
  }

  if (column.Title === column.StaticName) return;

  // Rename for display only. The internal name is already fixed by creation,
  // so this cannot break the item writes.
  await renameColumn(siteUrl, token, digest, title, column);
}

async function ensureColumns(siteUrl, token, digest, title, columns, onColumn) {
  const fields = await existingFields(siteUrl, token, title);

  for (const column of columns) {
    await ensureColumn(siteUrl, token, digest, title, column, fields);
    // Ticks on every path, including the skips, so the bar reflects progress
    // through the work rather than only through the columns that were missing.
    onColumn?.();
  }
}

/**
 * `onProgress(done, total)` counts columns checked across both lists. On a
 * first run this is around 70 sequential round trips and takes over a minute,
 * which looks identical to a hang unless something says otherwise.
 */
export async function provisionLists(siteUrl, token, { onProgress } = {}) {
  const digest = await getFormDigest(siteUrl, token);

  const total = DEVICE_COLUMNS.length + CHANGE_COLUMNS.length;
  let done = 0;
  const tick = () => {
    done += 1;
    onProgress?.(done, total);
  };

  await ensureList(
    siteUrl, token, digest, DEVICE_LIST_NAME,
    'One row per machine, from the scan reports',
  );
  await ensureColumns(siteUrl, token, digest, DEVICE_LIST_NAME, DEVICE_COLUMNS, tick);

  await ensureList(
    siteUrl, token, digest, CHANGE_LIST_NAME,
    'Field-level change history for the device list',
  );
  await ensureColumns(siteUrl, token, digest, CHANGE_LIST_NAME, CHANGE_COLUMNS, tick);

  return digest;
}

import { spFetch, getFormDigest, listPath } from './spClient.js';

/**
 * Making a SharePoint list look the way a section needs it: the list itself,
 * its columns, and the views that make those columns visible.
 *
 * This started life inside `devices/sharepoint/provisionLists.js` and was
 * lifted out whole when the asset register needed the same three steps against
 * a different schema. Everything device-specific stayed behind; what is here
 * takes the schema as an argument and knows nothing about what is in it.
 *
 * Every rule encoded below was verified against the tenant and each one fails
 * silently or confusingly if got wrong — see the comments at each site.
 */

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
 * A document library rather than a list: BaseTemplate 101, and it is where
 * uploaded photos land. Otherwise the same create-if-absent shape.
 */
async function ensureLibrary(siteUrl, token, digest, title, description) {
  const existing = await spFetch(siteUrl, listPath(title), { token });
  if (existing.ok) return;
  if (existing.status !== 404) {
    throw new Error(`Could not check for the "${title}" library (${existing.status})`);
  }

  const created = await spFetch(siteUrl, '/_api/web/lists', {
    token,
    digest,
    method: 'POST',
    body: {
      __metadata: { type: 'SP.List' },
      BaseTemplate: 101,
      Title: title,
      Description: description,
    },
  });

  if (!created.ok && created.status !== 201) {
    throw new Error(
      `Could not create the "${title}" library (${created.status}): ${await created.text()}`,
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
 * `defaultView` rather than a title lookup: the built-in view is called "All
 * Items" only in English, and this has to work on a site in any locale.
 */
const viewPath = (view) => (view.isDefault
  ? `${listPath(view.list)}/defaultView`
  : `${listPath(view.list)}/views/getByTitle('${encodeURIComponent(view.title)}')`);

async function viewTitles(siteUrl, token, listTitle) {
  const response = await spFetch(siteUrl, `${listPath(listTitle)}/views?$select=Title`, { token });
  if (!response.ok) {
    throw new Error(`Could not read the views of "${listTitle}" (${response.status})`);
  }

  const data = await response.json();
  return new Set((data.d?.results ?? []).map((v) => v.Title));
}

async function currentViewFields(siteUrl, token, view) {
  const response = await spFetch(siteUrl, `${viewPath(view)}/viewfields`, { token });
  if (!response.ok) return null;

  const data = await response.json();
  return data.d?.Items?.results ?? null;
}

async function currentViewQuery(siteUrl, token, view) {
  const response = await spFetch(siteUrl, `${viewPath(view)}?$select=ViewQuery`, { token });
  if (!response.ok) return null;

  const data = await response.json();
  return data.d?.ViewQuery ?? '';
}

/**
 * ViewQuery is only accepted in the creation body, and a default view is never
 * created -- so a filter or sort declared on one has to be merged on
 * afterwards or it is silently ignored.
 */
async function setViewQuery(siteUrl, token, digest, view) {
  const response = await spFetch(siteUrl, viewPath(view), {
    token,
    digest,
    method: 'POST',
    body: { __metadata: { type: 'SP.View' }, ViewQuery: view.query },
    headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
  });

  if (!response.ok) {
    throw new Error(
      `Could not set the "${view.title}" view query (${response.status}): `
        + `${await response.text()}`,
    );
  }
}

async function setViewFields(siteUrl, token, digest, view) {
  const base = `${viewPath(view)}/viewfields`;

  const cleared = await spFetch(siteUrl, `${base}/removeallviewfields`, {
    token, digest, method: 'POST',
  });
  if (!cleared.ok) {
    throw new Error(
      `Could not clear the "${view.title}" view (${cleared.status}): ${await cleared.text()}`,
    );
  }

  // One request per field, and order is the order they are added.
  for (const field of view.fields) {
    const added = await spFetch(siteUrl, `${base}/addviewfield('${field}')`, {
      token, digest, method: 'POST',
    });
    if (!added.ok) {
      throw new Error(
        `Could not add "${field}" to the "${view.title}" view (${added.status}): `
          + `${await added.text()}`,
      );
    }
  }
}

async function ensureView(siteUrl, token, digest, view, existingTitles) {
  if (!view.isDefault && !existingTitles.has(view.title)) {
    const created = await spFetch(siteUrl, `${listPath(view.list)}/views`, {
      token,
      digest,
      method: 'POST',
      body: {
        __metadata: { type: 'SP.View' },
        Title: view.title,
        ViewQuery: view.query ?? '',
        ViewType: 'HTML',
        RowLimit: 100,
        PersonalView: false,
      },
    });

    if (!created.ok && created.status !== 201) {
      throw new Error(
        `Could not create the "${view.title}" view (${created.status}): ${await created.text()}`,
      );
    }
  }

  if (view.query) {
    const currentQuery = await currentViewQuery(siteUrl, token, view);
    if (currentQuery !== view.query) await setViewQuery(siteUrl, token, digest, view);
  }

  // Rewriting the fields costs a request per field, so it only happens when
  // they are actually wrong -- otherwise every save would redo the whole set.
  const current = await currentViewFields(siteUrl, token, view);
  if (current && current.join('|') === view.fields.join('|')) return;

  await setViewFields(siteUrl, token, digest, view);
}

async function ensureViews(siteUrl, token, digest, views) {
  const byList = new Map();

  for (const view of views) {
    if (!byList.has(view.list)) {
      byList.set(view.list, await viewTitles(siteUrl, token, view.list));
    }
    await ensureView(siteUrl, token, digest, view, byList.get(view.list));
  }
}

/**
 * Bring a whole section's SharePoint schema to the shape it declares.
 *
 * `lists` is `[{ title, description, columns, library }]`; a `library` entry
 * gets a document library and no columns. `views` is the flat list of views
 * across all of them, and runs last because a view can only show columns that
 * already exist.
 *
 * `onProgress(done, total)` counts columns checked across every list. On a
 * first run this is dozens of sequential round trips and takes over a minute,
 * which looks identical to a hang unless something says otherwise.
 *
 * Returns the form digest, so a caller about to write rows does not have to
 * fetch a second one.
 */
export async function provisionSchema(siteUrl, token, { lists, views = [], onProgress } = {}) {
  const digest = await getFormDigest(siteUrl, token);

  const total = lists.reduce((sum, list) => sum + (list.columns?.length ?? 0), 0);
  let done = 0;
  const tick = () => {
    done += 1;
    onProgress?.(done, total);
  };

  for (const list of lists) {
    if (list.library) {
      await ensureLibrary(siteUrl, token, digest, list.title, list.description);
      continue;
    }

    // Provisioning runs to completion before any row is written: a half-created
    // list would fail every row with the same unhelpful message.
    await ensureList(siteUrl, token, digest, list.title, list.description);
    await ensureColumns(siteUrl, token, digest, list.title, list.columns, tick);
  }

  if (views.length) await ensureViews(siteUrl, token, digest, views);

  return digest;
}

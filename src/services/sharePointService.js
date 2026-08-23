/**
 * sharePointService.js
 * Idempotent SharePoint list + column provisioning using delegated user token.
 */

import { mergeChoices } from '../features/sharepoint/provision.js';

const LIST_NAME = 'IT Request Form';

// SP.FieldDateTime must be created via its own endpoint with correct __metadata type
// SP.Field (base) does NOT accept DisplayFormat — that's only on SP.FieldDateTime
const REQUIRED_COLUMNS = [
  { StaticName: 'Calling_x0020_Name', Title: 'Calling Name', FieldTypeKind: 2, spType: 'SP.Field' },
  { StaticName: 'Position', Title: 'Position', FieldTypeKind: 2, spType: 'SP.Field' },
  {
    StaticName: 'Department', Title: 'Department', FieldTypeKind: 6, spType: 'SP.FieldChoice',
  },
  {
    StaticName: 'Entity', Title: 'Entity', FieldTypeKind: 6, spType: 'SP.FieldChoice',
    // The reference form's entities. Reconciled onto the column ADDITIVELY by
    // `reconcileChoices` below, so `pmw`, `pmw-ss` and `pmw-th` stay valid on
    // every record already saved with them.
    choices: ['PMW', 'PCI', 'PML', 'WB']
  },
  { StaticName: 'Employee_x0020_ID', Title: 'Employee ID', FieldTypeKind: 2, spType: 'SP.Field' },
  { StaticName: 'Join_x0020__x002f__x0020_Last_x0', Title: 'Join / Last Work Date', FieldTypeKind: 4, spType: 'SP.FieldDate' },
  {
    StaticName: 'Equipment_x0020_Items', Title: 'Equipment Items', FieldTypeKind: 6, spType: 'SP.FieldMultiChoice',
    choices: ['laptop', 'pc', 'monitor', 'mouse', 'keyboard', 'headset']
  },
  { StaticName: 'Equipment_x0020_Remarks', Title: 'Equipment Remarks', FieldTypeKind: 3, spType: 'SP.Field' },
  {
    StaticName: 'Software_x0020_Licenses', Title: 'Software Licenses', FieldTypeKind: 6, spType: 'SP.FieldMultiChoice',
    choices: ['m365', 'sap', 'email']
  },
  { StaticName: 'Special_x0020_Permission', Title: 'Special Permission', FieldTypeKind: 3, spType: 'SP.Field' },
  {
    StaticName: 'Request_x0020_Type', Title: 'Request Type', FieldTypeKind: 6, spType: 'SP.FieldChoice',
    choices: ['Onboarding', 'Offboarding']
  },
];


// back to one header builder — verbose for everything
function buildHeaders(accessToken, formDigest = null) {
  const h = {
    'Accept': 'application/json;odata=verbose',
    'Content-Type': 'application/json;odata=verbose',
    'Authorization': `Bearer ${accessToken}`,
  };
  if (formDigest) h['X-RequestDigest'] = formDigest;
  return h;
}

function buildItemHeaders(accessToken, formDigest = null) {
  const h = {
    'Accept': 'application/json;odata=nometadata',
    'Content-Type': 'application/json;odata=nometadata',  // keep as-is
    'Authorization': `Bearer ${accessToken}`,
  };
  if (formDigest) h['X-RequestDigest'] = formDigest;
  return h;
}

async function getFormDigest(siteUrl, accessToken) {
  const res = await fetch(`${siteUrl}/_api/contextinfo`, {
    method: 'POST',
    headers: {
      'Accept': 'application/json;odata=verbose',
      'Content-Type': 'application/json;odata=verbose',
      'Authorization': `Bearer ${accessToken}`,
    },
  });
  if (!res.ok) throw new Error(`Failed to get form digest (${res.status})`);
  const data = await res.json();
  const digest = data?.d?.GetContextWebInformation?.FormDigestValue;
  if (!digest) throw new Error('Form digest value missing');
  return digest;
}

function toMultiChoice(values) {
  if (!values?.length) return null;
  return values; // ✅ plain array for odata=nometadata
}

async function ensureList(siteUrl, accessToken, formDigest) {
  const headers = buildHeaders(accessToken, formDigest);
  const listApiUrl = `${siteUrl}/_api/web/lists`;

  const checkRes = await fetch(
    `${listApiUrl}/getByTitle('${encodeURIComponent(LIST_NAME)}')`,
    { method: 'GET', headers }
  );

  if (checkRes.ok) {
    const data = await checkRes.json();
    console.log('[SP] List already exists:', LIST_NAME);
    return data.d;
  }

  if (checkRes.status !== 404) {
    const text = await checkRes.text();
    throw new Error(`Unexpected error checking list (${checkRes.status}): ${text}`);
  }

  console.log('[SP] Creating list:', LIST_NAME);
  const createRes = await fetch(listApiUrl, {
    method: 'POST',
    headers,
    body: JSON.stringify({
      __metadata: { type: 'SP.List' },
      BaseTemplate: 100,
      Title: LIST_NAME,
      Description: 'IT onboarding/offboarding request submissions',
    }),
  });

  if (!createRes.ok && createRes.status !== 201) {
    const text = await createRes.text();
    throw new Error(`Failed to create list (${createRes.status}): ${text}`);
  }

  const created = await createRes.json();
  console.log('[SP] List created:', created.d?.Id);
  return created.d;
}

async function getExistingFieldNames(siteUrl, accessToken) {
  const headers = buildHeaders(accessToken);
  const res = await fetch(
    `${siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(LIST_NAME)}')/fields?$select=StaticName`,
    { method: 'GET', headers }
  );
  if (!res.ok) throw new Error(`Failed to fetch fields (${res.status})`);
  const data = await res.json();
  return new Set((data.d?.results || []).map(f => f.StaticName));
}

/**
 * Bring the declared options onto a choice column that already exists.
 *
 * `ensureColumns` skips a column it finds, so an option added to the
 * declaration after the list was created never reaches SharePoint — invisible,
 * because the column is there and nothing complains until somebody cannot pick
 * the value they need.
 *
 * Additive only, using the same `mergeChoices` rule the shared provisioner
 * uses: a SharePoint row holding a value no longer in its column's list becomes
 * unreadable in that list, so an option is never taken away.
 */
async function reconcileChoices(siteUrl, accessToken, formDigest) {
  const listUrl = `${siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(LIST_NAME)}')`;

  for (const col of REQUIRED_COLUMNS) {
    if (!col.choices?.length) continue;

    const fieldUrl = `${listUrl}/fields/getByInternalNameOrTitle('${col.StaticName}')`;
    const read = await fetch(`${fieldUrl}?$select=Choices`, {
      method: 'GET',
      headers: buildHeaders(accessToken),
    });
    if (!read.ok) continue;

    const data = await read.json();
    const existing = data.d?.Choices?.results;
    if (!existing) continue;

    const merged = mergeChoices(existing, col.choices);
    if (merged.length === existing.length) continue;

    const res = await fetch(fieldUrl, {
      method: 'POST',
      headers: {
        ...buildHeaders(accessToken, formDigest),
        'X-HTTP-Method': 'MERGE',
        'IF-MATCH': '*',
      },
      body: JSON.stringify({
        // Never the base SP.Field: it does not declare `Choices`, and the
        // tenant answers "The property 'Choices' does not exist on type
        // 'SP.Field'".
        __metadata: { type: 'SP.FieldChoice' },
        Choices: { results: merged },
      }),
    });

    if (!res.ok) {
      console.warn('[SP] Could not update the options on', col.StaticName, res.status);
    }
  }
}

async function ensureColumns(siteUrl, accessToken, formDigest) {
  const headers = buildHeaders(accessToken, formDigest);
  const existingFields = await getExistingFieldNames(siteUrl, accessToken);
  const fieldsUrl = `${siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(LIST_NAME)}')/fields`;

  for (const col of REQUIRED_COLUMNS) {
    if (existingFields.has(col.StaticName)) {
      console.log('[SP] Column exists, skipping:', col.StaticName);
      continue;
    }

    console.log('[SP] Creating column:', col.StaticName);

    // ✅ FIX: DateTime fields MUST use SP.FieldDateTime as __metadata type.
    // SP.Field (the base type) does NOT have DisplayFormat — sending it causes 400.
    // For text/note fields use SP.Field. For date fields use SP.FieldDateTime.
    // For CREATION, always use SP.Field — SharePoint will set the correct type based on FieldTypeKind.
    const body = {
      __metadata: { type: col.spType === 'SP.FieldDateTime' ? 'SP.FieldDateTime' : 'SP.Field' },
      Title: col.Title,
      StaticName: col.StaticName,
      FieldTypeKind: col.FieldTypeKind,
      Required: false,
    };

    // DisplayFormat is valid ONLY on SP.FieldDateTime
    if (col.spType === 'SP.FieldDateTime') {
      body.DisplayFormat = 0; // 0 = DateOnly, 1 = DateTime
    }

    if (col.spType === 'SP.FieldMultiChoice') {
      body.AllowMultipleChoices = true;
    }

    const res = await fetch(fieldsUrl, {
      method: 'POST',
      headers,
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      if (res.status === 409) {
        console.warn('[SP] Column conflict (race condition), skipping:', col.StaticName);
        continue;
      }
      const text = await res.text();
      throw new Error(`Failed to create column "${col.StaticName}" (${res.status}): ${text}`);
    }

    console.log('[SP] Column created:', col.StaticName);
  }

  // Last, and over every choice column rather than only the ones just created:
  // the options a column offers can fall behind the declaration long after the
  // column itself exists.
  await reconcileChoices(siteUrl, accessToken, formDigest);
}

async function addListItem(siteUrl, accessToken, formDigest, itemData) {
  // Use nometadata for list items - simpler, no __metadata needed
  const headers = buildItemHeaders(accessToken, formDigest);
  const res = await fetch(
    `${siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(LIST_NAME)}')/items`,
    {
      method: 'POST',
      headers,
      body: JSON.stringify(itemData),
    }
  );
  if (!res.ok) throw new Error(`Failed to add list item (${res.status}): ${await res.text()}`);
  return await res.json();
}

// The asset checklist moved to `src/features/forms/sharepoint/`, onto the
// shared provisioner. What was here created its columns with `Choices` on a
// base `SP.Field` — a shape the tenant rejects, and which worked only because
// the list predated the bug; on a fresh site it would have failed outright.

export async function submitEmployeesToSharePoint(siteUrl, accessToken, employees, requestType) {
  if (!accessToken) throw new Error('No access token');
  if (!siteUrl) throw new Error('SharePoint site URL not configured');
  if (!employees?.length) throw new Error('No employee data to submit');

  const formDigest = await getFormDigest(siteUrl, accessToken);
  await ensureList(siteUrl, accessToken, formDigest);
  await ensureColumns(siteUrl, accessToken, formDigest);

  const results = [];

  for (const emp of employees) {
    // Fix: Ensure Join Date is in ISO 8601 format for SharePoint DateTime field
    let joinDateFormatted = null;
    if (emp.joinDate) {
      const d = new Date(emp.joinDate);
      if (!isNaN(d.getTime())) {
        joinDateFormatted = d.toISOString();
      }
    }

    // Build itemData, only including fields with actual values
    const itemData = {
      Title: emp.fullName || '',
      Calling_x0020_Name: emp.callingName || '',
      Position: emp.position || '',
      Equipment_x0020_Remarks: emp.equipmentRemarks || '',
      Employee_x0020_ID: emp.employeeId || '',
      Special_x0020_Permission: emp.specialPermission || '',
    };

    // Only add fields that have values (SharePoint rejects null for some types)
    if (emp.entity) itemData.Entity = emp.entity;
    if (emp.department) itemData.Department = emp.department;
    if (joinDateFormatted) itemData.Join_x0020__x002f__x0020_Last_x0 = joinDateFormatted;
    if (emp.equipmentItems?.length) itemData.Equipment_x0020_Items = toMultiChoice(emp.equipmentItems);
    if (emp.softwareLicenses?.length) itemData.Software_x0020_Licenses = toMultiChoice(emp.softwareLicenses);
    itemData.Request_x0020_Type = requestType;

    console.log('[SP] itemData:', JSON.stringify(itemData, null, 2));

    const result = await addListItem(siteUrl, accessToken, formDigest, itemData);
    results.push(result);
    console.log('[SP] Item submitted for:', emp.fullName);
  }

  return results;
}

// ─── Fetch column choices from SharePoint ─────────────────────────────────────

/**
 * Fetches the allowed choices for a specific column from the SharePoint list.
 * Works for Choice and MultiChoice field types.
 *
 * @param {string} siteUrl      - SharePoint site URL
 * @param {string} accessToken  - delegated user token
 * @param {string} staticName   - StaticName of the column (e.g. 'Entity')
 * @returns {Promise<string[]>} - array of choice strings
 */
export async function fetchColumnChoices(siteUrl, accessToken, staticName) {
  const headers = buildHeaders(accessToken);
  const res = await fetch(
    `${siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(LIST_NAME)}')/fields?$filter=StaticName eq '${staticName}'&$select=Choices,Title,StaticName,Id`,
    { method: 'GET', headers }
  );
  if (!res.ok) throw new Error(`Failed to fetch choices for ${staticName} (${res.status})`);
  const data = await res.json();
  const results = data?.d?.results || [];
  console.log(`[SP] field info for ${staticName}:`, JSON.stringify(results[0], null, 2)); // ← add this
  return results[0]?.Choices?.results || [];
}

/**
 * Fetches choices for multiple columns in parallel.
 *
 * @param {string}   siteUrl     - SharePoint site URL
 * @param {string}   accessToken - delegated user token
 * @param {string[]} staticNames - array of column StaticNames to fetch
 * @returns {Promise<Record<string, string[]>>} - map of StaticName → choices[]
 */
export async function fetchAllColumnChoices(siteUrl, accessToken, staticNames) {
  const entries = await Promise.all(
    staticNames.map(async (name) => {
      const choices = await fetchColumnChoices(siteUrl, accessToken, name);
      return [name, choices];
    })
  );
  return Object.fromEntries(entries);
}

// ─── Fetch all list items ───────────────────────────────────────────────────────────

export async function fetchAllListItems(siteUrl, accessToken) {
  const headers = buildHeaders(accessToken);
  const res = await fetch(
    `${siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(LIST_NAME)}')/items?$orderby=ID desc&$top=100`,
    { method: 'GET', headers }
  );
  if (!res.ok) throw new Error(`Failed to fetch list items (${res.status})`);
  const data = await res.json();
  return data.d?.results || [];
}

// ─── Update a list item ──────────────────────────────────────────────────

export async function updateListItem(siteUrl, accessToken, itemId, itemData) {
  const headers = {
    'Accept': 'application/json;odata=verbose',
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${accessToken}`,
  };
  const formDigest = await getFormDigest(siteUrl, accessToken);
  const res = await fetch(
    `${siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(LIST_NAME)}')/items(${itemId})`,
    {
      method: 'PATCH',
      headers: { ...headers, 'X-RequestDigest': formDigest, 'IF-MATCH': '*' },
      body: JSON.stringify(itemData),
    }
  );
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Failed to update list item (${res.status}): ${text}`);
  }
  // PATCH returns 204 No Content on success
  return { success: true };
}

// ─── Fetch single list item ───────────────────────────────────────────────

export async function fetchListItemById(siteUrl, accessToken, itemId) {
  const headers = buildHeaders(accessToken);
  const res = await fetch(
    `${siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(LIST_NAME)}')/items(${itemId})`,
    { method: 'GET', headers }
  );
  if (!res.ok) throw new Error(`Failed to fetch list item (${res.status})`);
  const data = await res.json();
  return data.d;
}

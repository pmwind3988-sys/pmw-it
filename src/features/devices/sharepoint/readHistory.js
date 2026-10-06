import { spFetch, listPath } from '../../sharepoint/spClient.js';
import { CHANGE_LIST_NAME } from './deviceSchema.js';

/**
 * One machine's change rows. Rows written since 2026-10-05 carry DeviceId;
 * older rows only the computer name, so both are asked for. A machine renamed
 * BEFORE DeviceId existed loses its pre-rename rows from this view -- they are
 * still in the list, under the old name.
 *
 * The filtered read is tried first. SharePoint refuses it once the list passes
 * the 5,000-item view threshold and the filter columns are not indexed -- that
 * comes back as a 500, not a 4xx -- and refuses it with a 400 while DeviceId
 * has not been provisioned. Either way the whole list is read instead, paged by
 * ID (which no threshold blocks), and matched here. Indexing DeviceId and
 * Title on IT Device Changes makes the first read succeed again.
 */
export async function readChanges(siteUrl, token, device) {
  const name = String(device.computerName ?? '').replace(/'/g, "''");
  const items = `${siteUrl}${listPath(CHANGE_LIST_NAME)}/items`;
  const filter = encodeURIComponent(`DeviceId eq ${Number(device.id)} or Title eq '${name}'`);

  let rows = await readPages(`${items}?$top=500&$filter=${filter}`, token);
  if (rows === null) return [];
  if (rows.refused) {
    rows = await readPages(`${items}?$top=5000`, token);
    if (rows === null) return [];
    if (rows.refused) throw new Error(`Could not read the change history (${rows.status})`);
  }

  // A reused name must not pull in another machine's rows: Title decides only
  // for rows that carry no DeviceId.
  const mine = rows.filter((row) => {
    const rowId = row.DeviceId === null || row.DeviceId === undefined || row.DeviceId === '' ? null : Number(row.DeviceId);
    if (rowId !== null) return rowId === Number(device.id);
    return String(row.Title ?? '').toLowerCase() === String(device.computerName ?? '').toLowerCase();
  });

  return mine.map((row) => ({
    deviceId: row.DeviceId === null || row.DeviceId === undefined || row.DeviceId === '' ? null : Number(row.DeviceId),
    fieldName: row.FieldName,
    oldValue: row.OldValue ?? '',
    newValue: row.NewValue ?? '',
    changeType: row.ChangeType,
    changedBy: row.ChangedBy ?? '',
    changedOn: row.ChangedOn ? new Date(row.ChangedOn).getTime() : null,
  }));
}

/**
 * Every page of one query: the rows, null when the list does not exist yet,
 * or `{ refused, status }` when SharePoint turned the FIRST page down. A
 * failure on a later page is not a refusal of the query, so it throws.
 */
async function readPages(firstUrl, token) {
  const rows = [];
  let url = firstUrl;
  while (url) {
    const response = await spFetch('', url, { token });
    if (response.status === 404) return null;
    if (!response.ok) {
      if (rows.length === 0 && url === firstUrl) return { refused: true, status: response.status };
      throw new Error(`Could not read the change history (${response.status})`);
    }
    const data = await response.json();
    rows.push(...(data.d?.results ?? []));
    url = data.d?.__next ?? null;
  }
  return rows;
}

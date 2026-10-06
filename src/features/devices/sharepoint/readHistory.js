import { spFetch, listPath } from '../../sharepoint/spClient.js';
import { CHANGE_LIST_NAME } from './deviceSchema.js';

/**
 * One machine's change rows. Rows written since 2026-10-05 carry DeviceId;
 * older rows only the computer name, so both are asked for. A machine renamed
 * BEFORE DeviceId existed loses its pre-rename rows from this view -- they are
 * still in the list, under the old name.
 *
 * A filter on a non-indexed column fails once the list passes SharePoint's
 * 5,000-item view threshold. If that day comes, index DeviceId and Title on
 * IT Device Changes in the list settings.
 */
export async function readChanges(siteUrl, token, device) {
  const name = String(device.computerName ?? '').replace(/'/g, "''");
  const base = `${siteUrl}${listPath(CHANGE_LIST_NAME)}/items?$top=500&$filter=`;
  const withId = encodeURIComponent(`DeviceId eq ${Number(device.id)} or Title eq '${name}'`);
  const titleOnly = encodeURIComponent(`Title eq '${name}'`);
  let url = `${base}${withId}`;
  let retriedWithoutId = false;
  const rows = [];

  while (url) {
    const response = await spFetch('', url, { token });
    if (response.status === 404) return [];
    // 400 on the first page: DeviceId has not been provisioned yet.
    if (response.status === 400 && !retriedWithoutId && rows.length === 0) {
      retriedWithoutId = true;
      url = `${base}${titleOnly}`;
      continue;
    }
    if (!response.ok) throw new Error(`Could not read the change history (${response.status})`);
    const data = await response.json();
    rows.push(...(data.d?.results ?? []));
    url = data.d?.__next ?? null;
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

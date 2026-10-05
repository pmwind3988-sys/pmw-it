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
  const filter = encodeURIComponent(`DeviceId eq ${Number(device.id)} or Title eq '${name}'`);
  let url = `${siteUrl}${listPath(CHANGE_LIST_NAME)}/items?$top=500&$filter=${filter}`;
  const rows = [];

  while (url) {
    const response = await spFetch('', url, { token });
    if (response.status === 404) return [];
    if (!response.ok) throw new Error(`Could not read the change history (${response.status})`);
    const data = await response.json();
    rows.push(...(data.d?.results ?? []));
    url = data.d?.__next ?? null;
  }

  return rows.map((row) => ({
    fieldName: row.FieldName,
    oldValue: row.OldValue ?? '',
    newValue: row.NewValue ?? '',
    changeType: row.ChangeType,
    changedBy: row.ChangedBy ?? '',
    changedOn: row.ChangedOn ? new Date(row.ChangedOn).getTime() : null,
  }));
}

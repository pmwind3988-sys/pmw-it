import { spFetch, listPath } from '../../sharepoint/spClient.js';
import { ASSIGNMENT_LIST_NAME, fromAssignmentItem } from './assignmentSchema.js';

const PAGE_SIZE = 500;

/**
 * Every stint, or one machine's. A list that does not exist yet is the state
 * before the first lifecycle write and reads as no history, not as a failure.
 */
export async function readAssignments(siteUrl, token, { deviceId } = {}) {
  const filter = deviceId === undefined || deviceId === null
    ? ''
    : `&$filter=${encodeURIComponent(`DeviceId eq ${Number(deviceId)}`)}`;
  let url = `${siteUrl}${listPath(ASSIGNMENT_LIST_NAME)}/items?$top=${PAGE_SIZE}${filter}`;
  const rows = [];

  while (url) {
    const response = await spFetch('', url, { token });
    if (response.status === 404) return [];
    if (!response.ok) throw new Error(`Could not read the owner history (${response.status})`);
    const data = await response.json();
    rows.push(...(data.d?.results ?? []));
    url = data.d?.__next ?? null;
  }

  return rows.map(fromAssignmentItem);
}

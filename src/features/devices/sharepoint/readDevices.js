import { spFetch, listPath } from '../../sharepoint/spClient.js';
import { DEVICE_LIST_NAME, fromListItem } from './deviceSchema.js';
import { refixStored } from '../derive/refixStored.js';

const PAGE_SIZE = 500;

/**
 * Reads every row once so the whole batch can be diffed in memory. The
 * alternative — one `$filter=Title eq '…'` per dropped file — is one request
 * per file and needs Title indexed to survive the 5,000-item view threshold.
 */
export async function readAllDevices(siteUrl, token) {
  const rows = [];
  let url = `${siteUrl}${listPath(DEVICE_LIST_NAME)}/items?$top=${PAGE_SIZE}`;

  while (url) {
    // `d.__next` is an absolute URL, so after the first page the address is
    // already complete and siteUrl must not be prefixed again.
    const response = await spFetch('', url, { token });

    // The list not existing yet is not an error — it is the state before the
    // first import, and the caller wants an empty register, not a failure.
    if (response.status === 404) return [];
    if (!response.ok) throw new Error(`Could not read the device list (${response.status})`);

    const data = await response.json();
    rows.push(...(data.d?.results ?? []));
    url = data.d?.__next ?? null;
  }

  // Each row is brought in line with today's exclusion rules as it is read, so
  // a machine scanned before its IT extraction disk was known stops reporting
  // that disk as its own storage without waiting to be scanned again.
  return rows.map(fromListItem).map(refixStored);
}

/**
 * One machine as SharePoint holds it NOW. A lifecycle action re-reads before
 * it plans, the same rule as a handover: two people moving the same laptop
 * from two screens must not both succeed.
 */
export async function readDevice(siteUrl, token, id) {
  const response = await spFetch(siteUrl, `${listPath(DEVICE_LIST_NAME)}/items(${Number(id)})`, { token });
  if (response.status === 404) return null;
  if (!response.ok) throw new Error(`Could not read that machine (${response.status})`);
  const data = await response.json();
  return refixStored(fromListItem(data.d ?? data));
}

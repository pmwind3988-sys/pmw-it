import { spFetch, listPath } from '../../sharepoint/spClient.js';
import { ASSET_LIST_NAME, fromListItem } from './assetSchema.js';

const PAGE_SIZE = 500;

/**
 * Reads every row once so a whole delivery can be planned in memory. The
 * alternative — one `$filter=AssetKey eq '…'` per scanned item — is a request
 * per item and needs the column indexed to survive the 5,000-item threshold.
 */
export async function readAllAssets(siteUrl, token) {
  const rows = [];
  let url = `${siteUrl}${listPath(ASSET_LIST_NAME)}/items?$top=${PAGE_SIZE}`;

  while (url) {
    // `d.__next` is an absolute URL, so after the first page the address is
    // already complete and siteUrl must not be prefixed again.
    const response = await spFetch('', url, { token });

    // The list not existing yet is not an error — it is the state before the
    // first save, and the caller wants an empty register, not a failure.
    if (response.status === 404) return [];
    if (!response.ok) throw new Error(`Could not read the asset register (${response.status})`);

    const data = await response.json();
    rows.push(...(data.d?.results ?? []));
    url = data.d?.__next ?? null;
  }

  return rows.map(fromListItem);
}

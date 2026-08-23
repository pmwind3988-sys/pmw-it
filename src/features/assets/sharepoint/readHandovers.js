import { spFetch, listPath } from '../../sharepoint/spClient.js';
import { HANDOVER_LIST_NAME, fromListItem } from './handoverSchema.js';

const PAGE_SIZE = 500;

/**
 * Every handover, read once. The people list, the person page, an item's
 * history and the overdue figure all come off this single read, so a count on a
 * card and the rows it opens cannot disagree.
 */
export async function readAllHandovers(siteUrl, token) {
  const rows = [];
  let url = `${siteUrl}${listPath(HANDOVER_LIST_NAME)}/items?$top=${PAGE_SIZE}`;

  while (url) {
    // `d.__next` is absolute, so after the first page siteUrl must not be
    // prefixed again.
    const response = await spFetch('', url, { token });

    // The list not existing yet is the state before the first handover, not an
    // error — the caller wants an empty history, not a failure.
    if (response.status === 404) return [];
    if (!response.ok) throw new Error(`Could not read the handovers (${response.status})`);

    const data = await response.json();
    rows.push(...(data.d?.results ?? []));
    url = data.d?.__next ?? null;
  }

  return rows.map(fromListItem);
}

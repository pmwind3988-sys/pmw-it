import { spFetch, listPath, ITEM_ACCEPT } from '../../sharepoint/spClient.js';
import { withRetry } from '../../sharepoint/writePool.js';
import { provisionSchema } from '../../sharepoint/provision.js';
import { STANDARDS_LIST_NAME, STANDARDS_COLUMNS, toStandardItem } from './standardsSchema.js';
import { validateStandard } from '../standards/validateStandard.js';

/** Append a version. Never updates a row: the list is the history. */
export async function saveStandard({
  siteUrl, token, standard, version, summary, savedBy, now = Date.now(),
}) {
  const check = validateStandard(standard);
  if (!check.ok) throw new Error(check.errors[0].message);

  const digest = await provisionSchema(siteUrl, token, {
    lists: [{
      title: STANDARDS_LIST_NAME,
      description: 'Versions of what counts as Critical, Needs attention, Moderate and Optimal for each part',
      columns: STANDARDS_COLUMNS,
    }],
  });

  const response = await withRetry(() => spFetch(siteUrl, `${listPath(STANDARDS_LIST_NAME)}/items`, {
    token, digest, method: 'POST', accept: ITEM_ACCEPT,
    body: toStandardItem({ version, standard, savedBy, summary, savedOn: now }),
  }));
  if (response.status === 403) throw new Error('You do not have edit rights on the IT Device Standards list. Ask IT.');
  if (!response.ok) throw new Error(`Could not save the standard (${response.status})`);
}

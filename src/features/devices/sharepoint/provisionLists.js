import { provisionSchema, fieldBody, renameBody } from '../../sharepoint/provision.js';
import {
  DEVICE_COLUMNS, CHANGE_COLUMNS, DEVICE_LIST_NAME, CHANGE_LIST_NAME,
} from './deviceSchema.js';
import { DEVICE_VIEWS } from './deviceViews.js';

/**
 * The generic engine — every SharePoint rule that took a day to learn — now
 * lives in `features/sharepoint/provision.js`, because the asset register needs
 * the same three steps against a different schema. What is left here is the
 * device section's declaration of what it wants.
 *
 * Re-exported so that `provisionLists.js` remains the one import for anything
 * doing device provisioning, tests included.
 */
export { fieldBody, renameBody };

/**
 * `onProgress(done, total)` counts columns checked across both lists. On a
 * first run this is around 70 sequential round trips and takes over a minute,
 * which looks identical to a hang unless something says otherwise.
 */
export function provisionLists(siteUrl, token, { onProgress } = {}) {
  return provisionSchema(siteUrl, token, {
    lists: [
      {
        title: DEVICE_LIST_NAME,
        description: 'One row per machine, from the scan reports',
        columns: DEVICE_COLUMNS,
      },
      {
        title: CHANGE_LIST_NAME,
        description: 'Field-level change history for the device list',
        columns: CHANGE_COLUMNS,
      },
    ],
    views: DEVICE_VIEWS,
    onProgress,
  });
}

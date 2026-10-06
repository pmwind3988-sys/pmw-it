import { provisionSchema, fieldBody, renameBody } from '../../sharepoint/provision.js';
import {
  DEVICE_COLUMNS, CHANGE_COLUMNS, DEVICE_LIST_NAME, CHANGE_LIST_NAME,
} from './deviceSchema.js';
import { ASSIGNMENT_LIST_NAME, ASSIGNMENT_COLUMNS } from './assignmentSchema.js';
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
 * `onProgress(done, total)` counts columns checked across all three lists. On a
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
      {
        title: ASSIGNMENT_LIST_NAME,
        description: 'Who had each machine, from when to when',
        columns: ASSIGNMENT_COLUMNS,
      },
    ],
    views: DEVICE_VIEWS,
    onProgress,
  });
}

/**
 * Only the columns of the device list and its change log -- no owner-history
 * list, no views. What an edit from a machine page needs when SharePoint has
 * just said a column is missing: one request per list when everything is
 * already there, against the dozens a full run spends checking views.
 */
export function provisionDeviceColumns(siteUrl, token) {
  return provisionSchema(siteUrl, token, {
    lists: [
      { title: DEVICE_LIST_NAME, description: 'One row per machine, from the scan reports', columns: DEVICE_COLUMNS },
      { title: CHANGE_LIST_NAME, description: 'Field-level change history for the device list', columns: CHANGE_COLUMNS },
    ],
  });
}

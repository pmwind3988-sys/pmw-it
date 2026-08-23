import { provisionSchema } from '../../sharepoint/provision.js';
import {
  ASSET_COLUMNS, BATCH_COLUMNS, CHANGE_COLUMNS,
  ASSET_LIST_NAME, BATCH_LIST_NAME, CHANGE_LIST_NAME, PHOTO_LIBRARY_NAME,
} from './assetSchema.js';
import { ASSET_VIEWS } from './assetViews.js';

/**
 * What the asset register needs to exist in SharePoint. Every rule about HOW
 * to create it lives in `features/sharepoint/provision.js`; this is only the
 * declaration.
 *
 * `onProgress(done, total)` counts columns across the three lists. On a first
 * run that is around fifty sequential round trips and takes over a minute,
 * which looks identical to a hang unless something says otherwise.
 */
export function provisionAssets(siteUrl, token, { onProgress } = {}) {
  return provisionSchema(siteUrl, token, {
    lists: [
      {
        title: ASSET_LIST_NAME,
        description: 'Everything IT owns: one row per tracked unit, one per bulk line',
        columns: ASSET_COLUMNS,
      },
      {
        title: BATCH_LIST_NAME,
        description: 'One row per delivery, holding the purchase details its items share',
        columns: BATCH_COLUMNS,
      },
      {
        title: CHANGE_LIST_NAME,
        description: 'Field-level change history for the asset register',
        columns: CHANGE_COLUMNS,
      },
      {
        title: PHOTO_LIBRARY_NAME,
        description: 'Item photographs and scanned purchase orders',
        library: true,
      },
    ],
    views: ASSET_VIEWS,
    onProgress,
  });
}

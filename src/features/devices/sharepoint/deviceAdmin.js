import { spFetch } from '../../sharepoint/spClient.js';
import { hasPermission } from './readStandards.js';

/**
 * SP.PermissionKind 26 = ManagePermissions. Only Full Control carries it --
 * Edit and Contribute (the Members group) do not -- so it reads as "an owner
 * of the IT helpdesk site", which is who the portal calls an admin.
 */
const MANAGE_PERMISSIONS = 26;

const unwrap = (data) => data?.d?.EffectiveBasePermissions ?? data?.d ?? data;

/**
 * May the signed-in user remove a machine from the register? SharePoint
 * decides, as it does for the grading standard. Anything that is not a clear
 * yes is a no: a button that removes records is the one to hide when unsure.
 */
export async function isDeviceAdmin(siteUrl, token) {
  const response = await spFetch(siteUrl, '/_api/web/EffectiveBasePermissions', { token });
  if (!response.ok) return false;
  return hasPermission(unwrap(await response.json()), MANAGE_PERMISSIONS);
}

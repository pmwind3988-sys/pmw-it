import { spFetch, listPath } from '../../sharepoint/spClient.js';
import { STANDARDS_LIST_NAME, fromStandardItem } from './standardsSchema.js';
import { validateStandard } from '../standards/validateStandard.js';
import { DEFAULT_STANDARD } from '../standards/defaultStandard.js';

/**
 * The standard in force: the newest row that is valid. A broken newest row is
 * skipped AND named, never silently -- grading must not stop because somebody
 * saved something odd, and nobody should wonder why their change is not showing.
 */
export function pickStandard(rows) {
  const history = rows.map((row) => ({ ...row, valid: Boolean(row.standard) && validateStandard(row.standard).ok }));
  const inForce = history.find((row) => row.valid) ?? null;
  const skipped = history.length && history[0] !== inForce ? history[0] : null;
  return {
    standard: inForce ? inForce.standard : DEFAULT_STANDARD,
    version: inForce?.version ?? 0,
    savedBy: inForce?.savedBy ?? null,
    savedOn: inForce?.savedOn ?? null,
    history,
    note: skipped
      ? `Version ${skipped.version ?? '?'} could not be used, so ${inForce ? `version ${inForce.version}` : 'the default standard'} is in force.`
      : null,
  };
}

export async function readStandards(siteUrl, token) {
  const response = await spFetch(siteUrl, `${listPath(STANDARDS_LIST_NAME)}/items?$orderby=Id%20desc&$top=20`, { token });
  if (response.status === 404) return pickStandard([]);
  if (!response.ok) throw new Error(`Could not read the device standards (${response.status})`);
  const data = await response.json();
  return pickStandard((data.d?.results ?? []).map(fromStandardItem));
}

/** SP.PermissionKind values: 2 = AddListItems, 12 = ManageLists. */
const ADD_LIST_ITEMS = 2;
const MANAGE_LISTS = 12;

export function hasPermission(perms, kind) {
  const bit = kind - 1;
  const word = bit < 32 ? Number(perms?.Low ?? 0) : Number(perms?.High ?? 0);
  return ((word >>> (bit % 32)) & 1) === 1;
}

const unwrap = (data) => data?.d?.EffectiveBasePermissions ?? data?.d ?? data;

/**
 * May the signed-in user save a standard? SharePoint decides, not the portal:
 * edit rights on the list itself, or -- before the list exists -- the right to
 * create lists, which is what the first save needs.
 */
export async function canEditStandards(siteUrl, token) {
  const onList = await spFetch(siteUrl, `${listPath(STANDARDS_LIST_NAME)}/EffectiveBasePermissions`, { token });
  if (onList.ok) return hasPermission(unwrap(await onList.json()), ADD_LIST_ITEMS);
  if (onList.status !== 404) return false;
  const onSite = await spFetch(siteUrl, '/_api/web/EffectiveBasePermissions', { token });
  if (!onSite.ok) return false;
  return hasPermission(unwrap(await onSite.json()), MANAGE_LISTS);
}

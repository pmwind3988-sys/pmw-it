import { spFetch, listPath, getFormDigest } from '../../sharepoint/spClient.js';
import { withRetry } from '../../sharepoint/writePool.js';
import { mergeChoices } from '../../sharepoint/provision.js';
import { ASSET_LIST_NAME } from './assetSchema.js';
import { HANDOVER_LIST_NAME } from './handoverSchema.js';
import { cleanCategory } from '../categories.js';

/**
 * A new kind of thing, added to the register for everybody.
 *
 * `Category` is a SharePoint CHOICE column, which refuses a value that is not
 * one of its options. So a category typed on this phone is not usable until
 * the column itself has heard of it — otherwise the first save of an item
 * using it fails, with a message about a property on a list item type that
 * tells the person holding the phone nothing.
 *
 * BOTH lists are updated. The handover list carries a copy of the category on
 * every row it writes, and a handover of the first projector would fail on
 * exactly the same rule if only the register had been told.
 *
 * Options are only ever ADDED, never replaced — the same rule provisioning
 * keeps. A row saved under an option that later disappeared becomes unreadable
 * in its own list.
 */
const CATEGORY_FIELD = 'Category';

const fieldPath = (list) =>
  `${listPath(list)}/fields/getByInternalNameOrTitle('${CATEGORY_FIELD}')`;

async function optionsOn(siteUrl, token, list) {
  const response = await spFetch(siteUrl, `${fieldPath(list)}?$select=Choices`, { token });
  if (!response.ok) {
    throw new Error(`Could not read the categories on "${list}" (${response.status})`);
  }

  const data = await response.json();
  return data.d?.Choices?.results ?? null;
}

async function offer(siteUrl, token, digest, list, name) {
  const existing = await optionsOn(siteUrl, token, list);
  // Not knowing what the column offers is not a reason to rewrite it with a
  // guess — that would drop every option somebody added in SharePoint itself.
  if (!existing) return false;

  const merged = mergeChoices(existing, [name]);
  if (merged.length === existing.length) return false;

  const response = await withRetry(() => spFetch(siteUrl, fieldPath(list), {
    token,
    digest,
    method: 'POST',
    body: {
      __metadata: { type: 'SP.FieldChoice' },
      Choices: { results: merged },
    },
    headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
  }));

  if (!response.ok) {
    throw new Error(
      `Could not add "${name}" to the categories (${response.status}): ${await response.text()}`,
    );
  }

  return true;
}

export async function addCategory({ siteUrl, token, name }) {
  const category = cleanCategory(name);
  if (!category) throw new Error('A category needs a name');

  const digest = await getFormDigest(siteUrl, token);

  const added = await offer(siteUrl, token, digest, ASSET_LIST_NAME, category);
  // The register first and the handovers second, deliberately. If the second
  // fails, the category exists and can be recorded against an item; only
  // handing that item out would be refused, and running this again fixes it.
  const handovers = await offer(siteUrl, token, digest, HANDOVER_LIST_NAME, category);

  return { category, added: added || handovers };
}

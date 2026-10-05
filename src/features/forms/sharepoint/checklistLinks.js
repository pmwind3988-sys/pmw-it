import { spFetch, listPath, getFormDigest, ITEM_ACCEPT } from '../../sharepoint/spClient.js';
import { provisionSchema } from '../../sharepoint/provision.js';
import { withRetry } from '../../sharepoint/writePool.js';
import {
  CHECKLIST_LIST_NAME, SIGNATURE_LIBRARY_NAME, checklistColumns, CHECKLIST_VIEWS,
} from './checklistSchema.js';
import {
  LINKS_LIST_NAME, LINK_COLUMNS, LINK_VIEWS, LINK_STATUS, toLinkItem, fromLinkItem,
} from '../links/linkSchema.js';
import { newLinkCode } from '../links/linkCode.js';
import { expiryFields, reopenFields, planEdit } from '../links/linkChanges.js';
import { fileApiPath } from '../../assets/sharepoint/fileUrl.js';
import { cleanPreset, isBlankValue } from '../links/linkRules.js';
import { ENTITIES, fieldsFor } from '../checklistForm.js';

/**
 * Creating, listing and cancelling shared checklist links — from the portal,
 * as the signed-in person, over SharePoint REST like every other write here.
 *
 * Opening and submitting a link is the server's job (`server/`), because the
 * person doing it has no sign-in. Everything the server will need is
 * provisioned HERE, by somebody who can: the server's identity is only
 * granted write on the site, and provisioning lists is not its business.
 */

const DAY = 86400000;

/**
 * `entities` are the codes the link may offer. The checklist's Entity column
 * is a choice and SharePoint refuses a value it has never heard of — and the
 * server that will write the signed row has no right to add one — so every
 * code is merged in HERE, while somebody who can is creating the link.
 */
export function provisionLinks(siteUrl, token, { onProgress, entities = [] } = {}) {
  const known = [...new Set([...ENTITIES, ...entities])];
  return provisionSchema(siteUrl, token, {
    lists: [
      {
        title: LINKS_LIST_NAME,
        description: 'Shared asset checklist links — pre-filled by IT, signed by the employee',
        columns: LINK_COLUMNS,
      },
      {
        title: CHECKLIST_LIST_NAME,
        description: 'Signed asset checklists — what each employee received or handed back',
        columns: checklistColumns(known),
      },
      {
        title: SIGNATURE_LIBRARY_NAME,
        description: 'Signature images from the asset checklists',
        library: true,
      },
    ],
    views: [...LINK_VIEWS, ...CHECKLIST_VIEWS],
    onProgress,
  });
}

/**
 * The link as it will be stored. Pure, so "only a filled field can be locked
 * or opened, and only one this form type shows" is testable on its own.
 */
export function draftLink({
  formMode, values, editable = [], options = null,
  expiresInDays = 14, now = Date.now(), code = newLinkCode(),
}) {
  const preset = cleanPreset(values, options);
  const shown = new Set(fieldsFor(formMode));
  return {
    code,
    formMode,
    preset,
    // A blank field is the employee's anyway; listing it would say nothing.
    editable: editable.filter((field) => shown.has(field) && !isBlankValue(field, preset[field])),
    options,
    expiresOn: new Date(now + expiresInDays * DAY).toISOString(),
    status: LINK_STATUS.WAITING,
  };
}

export async function createLink({
  siteUrl, token, link, createdByName, createdByEmail, onProgress,
}) {
  const digest = await provisionLinks(siteUrl, token, {
    onProgress,
    entities: (link.options?.entities ?? []).map((entity) => entity.value),
  });

  const response = await withRetry(() => spFetch(siteUrl, `${listPath(LINKS_LIST_NAME)}/items`, {
    token,
    digest,
    method: 'POST',
    accept: ITEM_ACCEPT,
    body: toLinkItem(link, { createdByName, createdByEmail }),
  }));
  if (!response.ok) {
    throw new Error(`Could not create the link (${response.status}): ${await response.text()}`);
  }

  const data = await response.json();
  return { ...link, id: data.Id ?? data.ID ?? null };
}

/**
 * Newest first. A list nobody has created yet is an empty list, not an error.
 *
 * Every column comes back rather than a `$select`: naming a column the list
 * has not been given yet fails the whole read, and columns are added to this
 * list over time (the edit columns arrive the first time somebody edits).
 */
export async function listLinks({ siteUrl, token }) {
  const path = `${listPath(LINKS_LIST_NAME)}/items?$orderby=Created desc&$top=500`;
  const response = await spFetch(siteUrl, path, { token, accept: ITEM_ACCEPT });
  if (response.status === 404) return [];
  if (!response.ok) {
    throw new Error(`Could not load the shared checklists (${response.status}): ${await response.text()}`);
  }
  const data = await response.json();
  return (data.value ?? []).map(fromLinkItem);
}

export async function getLink({ siteUrl, token, id }) {
  const response = await spFetch(siteUrl, `${listPath(LINKS_LIST_NAME)}/items(${Number(id)})`, {
    token, accept: ITEM_ACCEPT,
  });
  if (response.status === 404) return null;
  if (!response.ok) {
    throw new Error(`Could not load the shared checklist (${response.status}): ${await response.text()}`);
  }
  return fromLinkItem(await response.json());
}

async function merge(siteUrl, token, digest, list, id, body, what) {
  const response = await withRetry(() => spFetch(siteUrl, `${listPath(list)}/items(${Number(id)})`, {
    token,
    digest,
    method: 'POST',
    accept: ITEM_ACCEPT,
    body,
    headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
  }));
  if (!response.ok) {
    throw new Error(`Could not ${what} (${response.status}): ${await response.text()}`);
  }
}

async function updateLinkRow({ siteUrl, token, id, fields, what }) {
  const digest = await getFormDigest(siteUrl, token);
  await merge(siteUrl, token, digest, LINKS_LIST_NAME, id, fields, what);
}

export function cancelLink({ siteUrl, token, id }) {
  return updateLinkRow({
    siteUrl, token, id, fields: { LinkStatus: LINK_STATUS.CANCELLED }, what: 'cancel the link',
  });
}

/** A new end date — also what "Expire now" is, with the date being now. */
export function setLinkExpiry({ siteUrl, token, id, expiresAt }) {
  return updateLinkRow({
    siteUrl, token, id, fields: expiryFields(expiresAt), what: 'change when the link expires',
  });
}

/** Back to waiting. A signed record and its signature stay until signed again. */
export function reopenLink({ siteUrl, token, link, expiresAt }) {
  return updateLinkRow({
    siteUrl, token, id: link.id, fields: reopenFields(link, expiresAt), what: 'reopen the link',
  });
}

/**
 * IT correcting a signed checklist. The row is written FIRST and the link
 * second: if the second write fails, the record already holds the correction
 * and says it was edited, which is the part people read.
 *
 * Provisioned first, because the edit columns arrive with the first edit and
 * the entity picked may be one the Entity choice column has not heard of.
 */
export async function editSignedLink({
  siteUrl, token, link, values, by, at = Date.now(),
}) {
  const plan = planEdit(link, values, { by, at });
  if (Object.keys(plan.errors).length) return { errors: plan.errors };

  const digest = await provisionLinks(siteUrl, token, {
    entities: [
      ...(link.options?.entities ?? []).map((entity) => entity.value),
      plan.values.entity,
    ].filter(Boolean),
  });
  await merge(siteUrl, token, digest, CHECKLIST_LIST_NAME, link.checklistId, plan.checklist, 'save the checklist');
  await merge(siteUrl, token, digest, LINKS_LIST_NAME, link.id, plan.link, 'record the edit on the link');
  return { errors: {}, values: plan.values };
}

// Recycled, not deleted: SharePoint keeps it in the site's recycle bin and it
// can be restored. Already gone counts as done, so a delete that failed
// halfway can simply be pressed again.
async function recycle(siteUrl, token, digest, path, what) {
  const response = await withRetry(() => spFetch(siteUrl, `${path}/recycle()`, {
    token, digest, method: 'POST', accept: ITEM_ACCEPT,
  }));
  if (!response.ok && response.status !== 404) {
    throw new Error(`Could not remove ${what} (${response.status}): ${await response.text()}`);
  }
}

/**
 * The link, the checklist it produced, and that checklist's signature — to
 * the recycle bin. The link goes LAST, so a failure part-way leaves it on the
 * list to be removed again rather than leaving an orphaned record nobody can
 * find from here.
 */
export async function deleteLink({ siteUrl, token, link }) {
  const digest = await getFormDigest(siteUrl, token);

  if (link.checklistId) {
    const rowPath = `${listPath(CHECKLIST_LIST_NAME)}/items(${Number(link.checklistId)})`;
    const response = await spFetch(siteUrl, `${rowPath}?$select=SignatureUrl`, { token, accept: ITEM_ACCEPT });
    if (response.ok) {
      const { SignatureUrl: signature } = await response.json();
      if (signature) {
        // The same file path the asset photos are read through, verified
        // against the tenant; without its `/$value` it addresses the file.
        const file = fileApiPath(signature).replace(/\/\$value$/, '');
        if (file) await recycle(siteUrl, token, digest, file, 'the signature');
      }
      await recycle(siteUrl, token, digest, rowPath, 'the signed checklist');
    } else if (response.status !== 404) {
      throw new Error(`Could not find the signed checklist (${response.status}): ${await response.text()}`);
    }
  }

  await recycle(siteUrl, token, digest, `${listPath(LINKS_LIST_NAME)}/items(${Number(link.id)})`, 'the link');
}

/** The address an employee opens. */
export const linkUrl = (origin, code) => `${String(origin).replace(/\/$/, '')}/c/${code}`;

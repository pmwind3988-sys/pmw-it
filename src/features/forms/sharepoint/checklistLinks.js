import { spFetch, listPath, getFormDigest, ITEM_ACCEPT } from '../../sharepoint/spClient.js';
import { provisionSchema } from '../../sharepoint/provision.js';
import { withRetry } from '../../sharepoint/writePool.js';
import {
  CHECKLIST_LIST_NAME, SIGNATURE_LIBRARY_NAME, CHECKLIST_COLUMNS, CHECKLIST_VIEWS,
} from './checklistSchema.js';
import {
  LINKS_LIST_NAME, LINK_COLUMNS, LINK_VIEWS, LINK_STATUS, toLinkItem, fromLinkItem,
} from '../links/linkSchema.js';
import { newLinkCode } from '../links/linkCode.js';
import { cleanPreset, isBlankValue } from '../links/linkRules.js';
import { fieldsFor } from '../checklistForm.js';

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

export function provisionLinks(siteUrl, token, { onProgress } = {}) {
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
        columns: CHECKLIST_COLUMNS,
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
export function draftLink({ formMode, values, editable = [], expiresInDays = 14, now = Date.now(), code = newLinkCode() }) {
  const preset = cleanPreset(values);
  const shown = new Set(fieldsFor(formMode));
  return {
    code,
    formMode,
    preset,
    // A blank field is the employee's anyway; listing it would say nothing.
    editable: editable.filter((field) => shown.has(field) && !isBlankValue(field, preset[field])),
    expiresOn: new Date(now + expiresInDays * DAY).toISOString(),
    status: LINK_STATUS.WAITING,
  };
}

export async function createLink({
  siteUrl, token, link, createdByName, createdByEmail, onProgress,
}) {
  const digest = await provisionLinks(siteUrl, token, { onProgress });

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

const SELECT = [
  'Id', 'Title', 'FormMode', 'EmployeeName', 'Editable', 'ExpiresOn', 'LinkStatus',
  'CreatedByName', 'CreatedByEmail', 'Created', 'Modified', 'ChecklistId', 'SignedOn',
].join(',');

/** Newest first. A list nobody has created yet is an empty list, not an error. */
export async function listLinks({ siteUrl, token }) {
  const path = `${listPath(LINKS_LIST_NAME)}/items?$select=${SELECT}&$orderby=Created desc&$top=500`;
  const response = await spFetch(siteUrl, path, { token, accept: ITEM_ACCEPT });
  if (response.status === 404) return [];
  if (!response.ok) {
    throw new Error(`Could not load the shared checklists (${response.status}): ${await response.text()}`);
  }
  const data = await response.json();
  return (data.value ?? []).map(fromLinkItem);
}

export async function cancelLink({ siteUrl, token, id }) {
  const digest = await getFormDigest(siteUrl, token);

  const response = await withRetry(() => spFetch(siteUrl, `${listPath(LINKS_LIST_NAME)}/items(${id})`, {
    token,
    digest,
    method: 'POST',
    accept: ITEM_ACCEPT,
    body: { LinkStatus: LINK_STATUS.CANCELLED },
    headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
  }));
  if (!response.ok) {
    throw new Error(`Could not cancel the link (${response.status}): ${await response.text()}`);
  }
}

/** The address an employee opens. */
export const linkUrl = (origin, code) => `${String(origin).replace(/\/$/, '')}/c/${code}`;

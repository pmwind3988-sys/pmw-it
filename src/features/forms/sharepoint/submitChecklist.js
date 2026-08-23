import { spFetch, spUpload, listPath, ITEM_ACCEPT } from '../../sharepoint/spClient.js';
import { provisionSchema } from '../../sharepoint/provision.js';
import { withRetry } from '../../sharepoint/writePool.js';
import {
  CHECKLIST_LIST_NAME, SIGNATURE_LIBRARY_NAME, CHECKLIST_COLUMNS, CHECKLIST_VIEWS,
} from './checklistSchema.js';
import { toChecklistItem, signatureFileName } from '../toChecklistItem.js';

/**
 * Sending a signed checklist to SharePoint.
 *
 * Replaces `submitAssetChecklistToSharePoint` in `sharePointService.js`, whose
 * `ensureAssetColumns` sent `Choices` on a base `SP.Field` — a shape the tenant
 * rejects, which worked only because its list predates the bug.
 */

export function provisionChecklist(siteUrl, token, { onProgress } = {}) {
  return provisionSchema(siteUrl, token, {
    lists: [
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
    views: CHECKLIST_VIEWS,
    onProgress,
  });
}

/** A data URL back to the bytes it encodes, without a fetch or a Blob. */
export function dataUrlToBytes(dataUrl) {
  const base64 = String(dataUrl ?? '').split(',')[1];
  if (!base64) return null;

  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

async function signatureFolder(siteUrl, token) {
  const path = `${listPath(SIGNATURE_LIBRARY_NAME)}/RootFolder?$select=ServerRelativeUrl`;
  const response = await spFetch(siteUrl, path, { token });
  if (!response.ok) throw new Error(`Could not find the signature library (${response.status})`);

  const data = await response.json();
  const url = data.d?.ServerRelativeUrl;
  if (!url) throw new Error('The signature library reported no folder');
  return url;
}

async function uploadSignature(siteUrl, token, digest, folder, dataUrl, fileName) {
  const bytes = dataUrlToBytes(dataUrl);
  if (!bytes) return '';

  const path = `/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(folder)}')`
    + `/Files/add(url='${encodeURIComponent(fileName)}',overwrite=true)`;

  const response = await spUpload(siteUrl, path, { token, digest, body: bytes });
  if (!response.ok) {
    throw new Error(`Could not upload the signature (${response.status}): ${await response.text()}`);
  }

  const data = await response.json();
  return data.d?.ServerRelativeUrl ?? `${folder}/${fileName}`;
}

/**
 * The signature goes up FIRST and a failure there fails the whole submission.
 *
 * That is deliberate and the opposite of how the asset register treats item
 * photos: a photo is an attachment to a record, but the signature IS the
 * record. A checklist row saved without one is not a weaker version of the
 * thing — it is a claim that somebody signed when they did not.
 */
export async function submitChecklist({
  siteUrl, token, values, onProgress,
}) {
  const report = (phase) => onProgress?.(phase);

  report('provisioning');
  const digest = await provisionChecklist(siteUrl, token, {
    onProgress: () => report('provisioning'),
  });

  report('signature');
  const submittedAt = Date.now();
  const folder = await signatureFolder(siteUrl, token);
  const signatureUrl = await uploadSignature(
    siteUrl, token, digest, folder,
    values.signature, signatureFileName(values, submittedAt),
  );

  report('saving');
  const response = await withRetry(() => spFetch(siteUrl, `${listPath(CHECKLIST_LIST_NAME)}/items`, {
    token,
    digest,
    method: 'POST',
    accept: ITEM_ACCEPT,
    body: toChecklistItem(values, { submittedAt, signatureUrl }),
  }));

  if (!response.ok) {
    throw new Error(`Could not save the checklist (${response.status}): ${await response.text()}`);
  }

  return { submittedAt, signatureUrl };
}

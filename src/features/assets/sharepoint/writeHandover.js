import { spFetch, listPath, ITEM_ACCEPT } from '../../sharepoint/spClient.js';
import { runPool, withRetry } from '../../sharepoint/writePool.js';
import { provisionAssets } from './provisionAssets.js';
import { readAllAssets } from './readAssets.js';
import { readAllHandovers } from './readHandovers.js';
import { ASSET_LIST_NAME, toUpdateItem as assetPatch } from './assetSchema.js';
import {
  HANDOVER_LIST_NAME, toListItem, toUpdateItem as handoverPatch,
} from './handoverSchema.js';
import { uploadSignature } from './uploadSignature.js';
import { planHandover } from '../handover/planHandover.js';
import { planReturn } from '../handover/planReturn.js';
import { planPersonEdit } from '../handover/planPersonEdit.js';

const itemPath = (listName) => `${listPath(listName)}/items`;

const insert = (siteUrl, token, digest, listName, body) => withRetry(
  () => spFetch(siteUrl, itemPath(listName), {
    token, digest, method: 'POST', body, accept: ITEM_ACCEPT,
  }),
);

/** A SharePoint update is a POST wearing these two headers. */
const merge = (siteUrl, token, digest, listName, id, body) => withRetry(
  () => spFetch(siteUrl, `${itemPath(listName)}(${id})`, {
    token,
    digest,
    method: 'POST',
    body,
    accept: ITEM_ACCEPT,
    headers: { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
  }),
);

/**
 * Handing a basket over.
 *
 * The register is re-read immediately before planning rather than trusted from
 * the screen. That is what stops two people issuing the same laptop from two
 * phones: the second one's plan sees the row already Assigned and refuses the
 * line (§8).
 */
export async function commitHandover({
  siteUrl, token, basket, issuedBy, signature, onProgress,
}) {
  const report = (phase, done = 0, total = 0) => onProgress?.({ phase, done, total });

  report('provisioning');
  const digest = await provisionAssets(siteUrl, token, {
    onProgress: (done, total) => report('provisioning', done, total),
  });

  report('reading');
  const register = await readAllAssets(siteUrl, token);

  const issuedOn = Date.now();

  // The signature is asked for and can be skipped, so one that will not upload
  // must not stop the handover: the laptop has changed hands either way, and a
  // register that refuses to say so is worse than one saying so unsigned.
  let issueSignature = '';
  let signatureFailed = '';
  if (signature) {
    report('signature');
    try {
      issueSignature = await uploadSignature({
        siteUrl,
        token,
        digest,
        dataUrl: signature,
        seed: basket.person?.email || basket.person?.name || 'handover',
      });
    } catch (thrown) {
      signatureFailed = thrown.message || 'The signature could not be saved';
    }
  }

  const plan = planHandover(basket, register, {
    issuedOn, issuedBy: issuedBy ?? '', issueSignature,
  });

  report('writing', 0, plan.handovers.length);
  const written = await runPool(plan.handovers, async (handover) => {
    const response = await insert(siteUrl, token, digest, HANDOVER_LIST_NAME, toListItem(handover));
    if (!response.ok) throw new Error(`${response.status}: ${await response.text()}`);
    return handover.assetKey;
  }, { concurrency: 4, onProgress: (done, total) => report('writing', done, total) });

  /**
   * The register copies go second, and only for the lines whose handover row
   * actually landed. The handover list is the truth (§4.2), so a register
   * update that fails leaves the item's copied fields stale rather than losing
   * the record — which is recoverable, where the reverse is not.
   */
  const landed = new Set(
    written.map((result, index) => (result.error ? null : plan.handovers[index].assetKey))
      .filter(Boolean),
  );

  const updates = plan.assetUpdates.filter((update) => landed.has(update.assetKey));
  report('updating', 0, updates.length);
  const updated = await runPool(updates, async (update) => {
    const response = await merge(
      siteUrl, token, digest, ASSET_LIST_NAME, update.id, assetPatch(update.body),
    );
    if (!response.ok) throw new Error(`${response.status}: ${await response.text()}`);
    return update.assetKey;
  }, { concurrency: 4, onProgress: (done, total) => report('updating', done, total) });

  return {
    handedOver: written.filter((result) => !result.error).length,
    signed: Boolean(issueSignature),
    signatureFailed,
    blocked: plan.blocked,
    writeFailures: written
      .map((result, index) => (result.error
        ? { assetKey: plan.handovers[index].assetKey, error: result.error.message }
        : null))
      .filter(Boolean),
    staleRows: updated
      .map((result, index) => (result.error ? updates[index].assetKey : null))
      .filter(Boolean),
  };
}

/**
 * Taking things back.
 *
 * Both lists are re-read for the same reason as above: somebody else may have
 * returned half of it from another screen, and returning three of a line with
 * two left has to be refused rather than driving `quantityOut` negative.
 */
export async function commitReturn({
  siteUrl, token, returns, returnedBy, signature, onProgress,
}) {
  const report = (phase, done = 0, total = 0) => onProgress?.({ phase, done, total });

  report('reading');
  const [register, handovers] = await Promise.all([
    readAllAssets(siteUrl, token),
    readAllHandovers(siteUrl, token),
  ]);

  const digest = await provisionAssets(siteUrl, token);

  // Same bargain as handing out: the thing is back on the shelf whether or not
  // the picture of the signature made it there.
  let returnSignature = '';
  let signatureFailed = '';
  if (signature) {
    report('signature');
    try {
      returnSignature = await uploadSignature({
        siteUrl, token, digest, dataUrl: signature, seed: 'return',
      });
    } catch (thrown) {
      signatureFailed = thrown.message || 'The signature could not be saved';
    }
  }

  const plan = planReturn(returns, handovers, register, {
    returnedOn: Date.now(),
    returnedBy: returnedBy ?? '',
    returnSignature,
  });

  report('writing', 0, plan.handoverUpdates.length);
  const written = await runPool(plan.handoverUpdates, async (update) => {
    const response = await merge(
      siteUrl, token, digest, HANDOVER_LIST_NAME, update.id, handoverPatch(update.body),
    );
    if (!response.ok) throw new Error(`${response.status}: ${await response.text()}`);
    return update.id;
  }, { concurrency: 4, onProgress: (done, total) => report('writing', done, total) });

  report('updating', 0, plan.assetUpdates.length);
  const updated = await runPool(plan.assetUpdates, async (update) => {
    const response = await merge(
      siteUrl, token, digest, ASSET_LIST_NAME, update.id, assetPatch(update.body),
    );
    if (!response.ok) throw new Error(`${response.status}: ${await response.text()}`);
    return update.assetKey;
  }, { concurrency: 4, onProgress: (done, total) => report('updating', done, total) });

  return {
    returned: written.filter((result) => !result.error).length,
    signed: Boolean(returnSignature),
    signatureFailed,
    blocked: plan.blocked,
    writeFailures: written.filter((result) => result.error).length,
    staleRows: updated
      .map((result, index) => (result.error ? plan.assetUpdates[index].assetKey : null))
      .filter(Boolean),
  };
}

/**
 * Correcting who somebody is.
 *
 * Both lists are re-read first, for the same reason the other two do it: the
 * rows being renamed may have changed since the page was opened, and a plan
 * built from a stale screen would rename a handover that has since been
 * returned by somebody else.
 *
 * Nothing here can move an item. The plan writes three fields on the handover
 * rows and two on the register rows that name this person, and nothing else —
 * so a correction cannot cost somebody the record of what they are holding.
 */
export async function commitPersonEdit({
  siteUrl, token, from, person, onProgress,
}) {
  const report = (phase, done = 0, total = 0) => onProgress?.({ phase, done, total });

  report('reading');
  const [register, handovers] = await Promise.all([
    readAllAssets(siteUrl, token),
    readAllHandovers(siteUrl, token),
  ]);

  const digest = await provisionAssets(siteUrl, token);

  const plan = planPersonEdit(handovers, register, {
    from,
    name: person.name,
    email: person.email,
    login: person.login,
  });

  report('writing', 0, plan.handoverUpdates.length);
  const written = await runPool(plan.handoverUpdates, async (update) => {
    const response = await merge(
      siteUrl, token, digest, HANDOVER_LIST_NAME, update.id, handoverPatch(update.body),
    );
    if (!response.ok) throw new Error(`${response.status}: ${await response.text()}`);
    return update.id;
  }, { concurrency: 4, onProgress: (done, total) => report('writing', done, total) });

  /**
   * The register copies go second and only where the handover row landed —
   * the handover list is the truth (§4.2), so a register copy left reading the
   * old name is stale rather than wrong, and recoverable by saving again.
   */
  report('updating', 0, plan.assetUpdates.length);
  const updated = await runPool(plan.assetUpdates, async (update) => {
    const response = await merge(
      siteUrl, token, digest, ASSET_LIST_NAME, update.id, assetPatch(update.body),
    );
    if (!response.ok) throw new Error(`${response.status}: ${await response.text()}`);
    return update.assetKey;
  }, { concurrency: 4, onProgress: (done, total) => report('updating', done, total) });

  return {
    rows: plan.rows,
    openLines: plan.openLines,
    changed: written.filter((result) => !result.error).length,
    writeFailures: written
      .map((result, index) => (result.error
        ? { id: plan.handoverUpdates[index].id, error: result.error.message }
        : null))
      .filter(Boolean),
    staleRows: updated
      .map((result, index) => (result.error ? plan.assetUpdates[index].assetKey : null))
      .filter(Boolean),
  };
}

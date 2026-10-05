import { isLinkCode } from '../src/features/forms/links/linkCode.js';
import {
  linkState, editableFields, mergeSubmission, cleanPreset,
} from '../src/features/forms/links/linkRules.js';
import { fromLinkItem, LINK_STATUS } from '../src/features/forms/links/linkSchema.js';
import { validateChecklist, hasErrors } from '../src/features/forms/validate.js';
import { toChecklistItem, signatureFileName } from '../src/features/forms/toChecklistItem.js';
import { dataUrlToBytes } from '../src/features/forms/sharepoint/submitChecklist.js';
import { ConflictError } from './errors.js';

/**
 * The two things an anonymous visitor can do with a link: open it, and submit
 * it once. Written over an injected Graph client so every rule here is tested
 * without a tenant (`checklistLinkApi.test.js`).
 *
 * Answers are `{ status, body }`; the Vercel entry turns them into HTTP.
 */

// Unknown, expired and cancelled all read the same, so a stranger trying codes
// learns nothing from the difference.
const GONE = { status: 404, body: { state: 'gone' } };

const SAVE_FAILED = 'Your form could not be saved just now. Your answers are still on the page — please try again.';

const bytesToDataUrl = (bytes) => `data:image/png;base64,${Buffer.from(bytes).toString('base64')}`;

export function createLinkApi({ graph, now = Date.now, log = console }) {
  async function load(code) {
    if (!isLinkCode(code)) return null;
    const row = await graph.findLink(code);
    if (!row) return null;
    return { row, link: fromLinkItem({ ...row.fields, id: row.id }) };
  }

  async function signedCopy(link) {
    let signature = null;
    if (link.signatureFile) {
      try {
        const bytes = await graph.readSignature(link.signatureFile);
        if (bytes) signature = bytesToDataUrl(bytes);
      } catch (error) {
        // The record still reads without its picture; say so in the log.
        log.warn?.('[checklist-link] signature could not be read', error);
      }
    }
    return {
      state: 'signed',
      formMode: link.formMode,
      values: link.submitted ?? {},
      signature,
      signedOn: link.signedOn,
    };
  }

  async function get(code) {
    const found = await load(code);
    if (!found) return GONE;
    const { link } = found;
    const state = linkState(link, now());

    if (state === 'signed') return { status: 200, body: await signedCopy(link) };
    if (state !== 'open' && state !== 'busy') return GONE;

    return {
      status: 200,
      body: {
        state: 'open',
        formMode: link.formMode,
        values: { formMode: link.formMode, ...cleanPreset(link.preset) },
        editable: editableFields(link),
        expiresOn: link.expiresOn,
      },
    };
  }

  async function submit(code, body) {
    const found = await load(code);
    if (!found) return GONE;
    const { row, link } = found;
    const state = linkState(link, now());

    if (state === 'signed') {
      return { status: 409, body: { state: 'signed', error: 'This form has already been signed.' } };
    }
    if (state === 'busy') {
      return { status: 409, body: { error: 'This form is being submitted already. Give it a moment and reload.' } };
    }
    if (state !== 'open') return GONE;

    const values = mergeSubmission(link, body?.values);
    const errors = validateChecklist(values);
    if (hasErrors(errors)) return { status: 422, body: { errors } };

    // Claimed before anything is written. The eTag is the one read above, so a
    // second submission that read the same row loses here and writes nothing.
    try {
      await graph.updateLink(row.id, { LinkStatus: LINK_STATUS.SIGNING }, { eTag: row.eTag });
    } catch (error) {
      if (error instanceof ConflictError || error?.name === 'ConflictError') {
        return { status: 409, body: { error: 'This form is being submitted already. Give it a moment and reload.' } };
      }
      log.error('[checklist-link] could not claim the link', error);
      return { status: 502, body: { error: SAVE_FAILED } };
    }

    const submittedAt = now();
    const signedOn = new Date(submittedAt).toISOString();
    const { signature, ...stored } = values;
    const signedCopyBody = {
      state: 'signed', formMode: link.formMode, values: stored, signature, signedOn,
    };
    let checklistId = null;
    try {
      const fileName = signatureFileName(values, submittedAt);
      const uploaded = await graph.uploadSignature(fileName, dataUrlToBytes(values.signature));
      const created = await graph.createChecklist(
        toChecklistItem(values, { submittedAt, signatureUrl: uploaded.serverRelativeUrl }),
      );
      checklistId = created.id;

      await markSigned(row.id, {
        LinkStatus: LINK_STATUS.SIGNED,
        Submitted: JSON.stringify(stored),
        SignatureFile: fileName,
        ChecklistId: Number(checklistId),
        SignedOn: signedOn,
      });

      return { status: 200, body: signedCopyBody };
    } catch (error) {
      log.error('[checklist-link] submission failed', { code, checklistId, error });
      // Only handed back if nothing was recorded. Once the checklist row
      // exists, reopening the link would let it be signed twice.
      if (checklistId === null) {
        await graph.updateLink(row.id, { LinkStatus: LINK_STATUS.WAITING }).catch((undo) => {
          log.error('[checklist-link] could not reopen the link', undo);
        });
        return { status: 502, body: { error: SAVE_FAILED } };
      }
      return { status: 200, body: signedCopyBody };
    }
  }

  // The checklist is already saved by now; a blip here must not lose the
  // record of where it went, so this one write is tried three times.
  async function markSigned(id, fields) {
    let last;
    for (let attempt = 0; attempt < 3; attempt += 1) {
      try {
        await graph.updateLink(id, fields);
        return;
      } catch (error) {
        last = error;
      }
    }
    throw last;
  }

  return { get, submit };
}

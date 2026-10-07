import { isLinkCode } from '../src/features/forms/links/linkCode.js';
import {
  linkState, editableFields, mergeSubmission, cleanPreset,
} from '../src/features/forms/links/linkRules.js';
import { fromLinkItem, LINK_STATUS } from '../src/features/forms/links/linkSchema.js';
import { isReopened } from '../src/features/forms/links/linkChanges.js';
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
  /**
   * A checklist made at the till is the signature for the handover (or the
   * return) it was made with: once signed, that signature goes onto those
   * rows, where the person page reads it. Never over one already there, and
   * never at the cost of the signing — the checklist is recorded either way,
   * and a row that could not be signed is logged, not failed.
   */
  async function attachToHandovers(link, signatureUrl) {
    const ids = Array.isArray(link.handovers?.ids) ? link.handovers.ids : [];
    if (!ids.length) return;
    const field = link.handovers.kind === 'return' ? 'ReturnSignature' : 'IssueSignature';
    for (const id of ids) {
      try {
        await graph.signHandover(id, field, signatureUrl);
      } catch (error) {
        log.warn?.('[checklist-link] signature not attached to handover row', { id, error });
      }
    }
  }

  async function load(code) {
    if (!isLinkCode(code)) return null;
    const row = await graph.findLink(code);
    if (!row) return null;
    return { row, link: fromLinkItem({ ...row.fields, id: row.id }) };
  }

  async function storedSignature(link) {
    if (!link.signatureFile) return null;
    try {
      const bytes = await graph.readSignature(link.signatureFile);
      return bytes ? bytesToDataUrl(bytes) : null;
    } catch (error) {
      // The record still reads without its picture; say so in the log.
      log.warn?.('[checklist-link] signature could not be read', error);
      return null;
    }
  }

  // Who edited a signed checklist, and when -- shown on the signed copy and
  // printed with it, because the signature no longer covers every value.
  const editedBy = (link) => (link.editedBy ? { by: link.editedBy, on: link.editedOn } : null);

  async function signedCopy(link) {
    return {
      state: 'signed',
      formMode: link.formMode,
      options: link.options,
      values: link.submitted ?? {},
      signature: await storedSignature(link),
      signedOn: link.signedOn,
      edited: editedBy(link),
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
        values: { formMode: link.formMode, ...cleanPreset(link.preset, link.options) },
        editable: editableFields(link),
        // The Entity and Department choices, as HR's lists stood when IT made
        // the link — this page has no way to read them itself.
        options: link.options,
        expiresOn: link.expiresOn,
        // A REOPENED link was signed before. The employee may keep that
        // signature rather than draw it again.
        existingSignature: isReopened(link) ? await storedSignature(link) : null,
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
    const reopened = isReopened(link);
    // Keeping a signature needs one to keep. A new one sent alongside wins.
    const keep = body?.keepSignature === true && reopened && Boolean(link.signatureFile)
      && !values.signature;
    const errors = validateChecklist(keep ? { ...values, signature: 'kept' } : values);
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
    const { signature: drawn, ...stored } = values;
    const signature = keep ? await storedSignature(link) : drawn;
    const signedCopyBody = {
      state: 'signed', formMode: link.formMode, options: link.options, values: stored, signature, signedOn, edited: null,
    };
    let written = false;
    let checklistId = link.checklistId;
    try {
      let fileName = link.signatureFile;
      let signatureUrl = null;
      if (!keep) {
        fileName = signatureFileName(values, submittedAt);
        // A new signature never overwrites the one it replaces.
        if (fileName === link.signatureFile) fileName = fileName.replace(/\.png$/, '-again.png');
        signatureUrl = (await graph.uploadSignature(fileName, dataUrlToBytes(drawn))).serverRelativeUrl;
      }

      const item = toChecklistItem(values, { submittedAt, signatureUrl: signatureUrl ?? '' });
      if (keep) delete item.SignatureUrl;
      // Signed again, so the signature covers the values as they now stand.
      if (link.editedBy) item.EditedAfterSigning = '';

      if (reopened) {
        // One record per link: a reopened link corrects the row it made.
        await graph.updateChecklist(link.checklistId, item);
      } else {
        checklistId = (await graph.createChecklist(item)).id;
      }
      written = true;

      await markSigned(row.id, {
        LinkStatus: LINK_STATUS.SIGNED,
        Submitted: JSON.stringify(stored),
        SignatureFile: fileName,
        ChecklistId: Number(checklistId),
        SignedOn: signedOn,
        ...(link.editedBy ? { EditedBy: '' } : {}),
      });

      if (signatureUrl) await attachToHandovers(link, signatureUrl);

      return { status: 200, body: signedCopyBody };
    } catch (error) {
      log.error('[checklist-link] submission failed', { code, checklistId, error });
      // Only handed back if nothing was recorded. Once the checklist row has
      // been written, reopening the link would let it be signed twice.
      if (!written) {
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

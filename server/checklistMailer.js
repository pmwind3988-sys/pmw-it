import { buildChecklistEmail, buildTestEmail, SIGNATURE_CID } from './checklistEmail.js';
import { EMAIL_SETTING_KEY, normaliseEmailSettings } from '../src/features/forms/links/emailSettings.js';
import { appendEmailLog, serialiseEmailLog, EMAIL_KINDS } from '../src/features/forms/links/emailLog.js';
import { fromLinkItem, LINK_STATUS } from '../src/features/forms/links/linkSchema.js';

/**
 * Emails about signed checklists, over an injected Graph client so every
 * rule is tested without a mailbox (`checklistMailer.test.js`).
 *
 * Three ways in:
 *   onSigned  the server, the moment an employee signs. Only while emails are
 *             ON, which is what makes "only new signings" true: a checklist
 *             signed before then never passed through here.
 *   resend    somebody in IT, for one signed checklist — old or new. Allowed
 *             while emails are off too, as long as a group is saved.
 *   test      somebody in IT, to see the group receive something.
 *
 * Recipients ALWAYS come from the saved setting, never from the request.
 * Every attempt, sent or failed, is written to the link's `EmailLog`; a
 * failure to send or to log never undoes a signing.
 */
export function createChecklistMailer({ graph, from, now = Date.now, log = console }) {
  async function settings() {
    try {
      return normaliseEmailSettings(await graph.readSetting(EMAIL_SETTING_KEY));
    } catch (error) {
      log.warn?.('[checklist-email] settings could not be read', error);
      return normaliseEmailSettings(null);
    }
  }

  async function send({ link, values, signedOn, origin, signatureBytes, resent, to }) {
    const hasSignature = Boolean(signatureBytes?.length);
    const mail = buildChecklistEmail({
      link, values, signedOn, origin, hasSignature, resent, to,
    });
    await graph.sendMail({
      from,
      to,
      subject: mail.subject,
      html: mail.html,
      attachments: hasSignature
        ? [{ name: 'signature.png', contentType: 'image/png', bytes: signatureBytes, contentId: SIGNATURE_CID }]
        : [],
    });
  }

  async function record(rowId, link, entry) {
    try {
      await graph.updateLink(rowId, { EmailLog: serialiseEmailLog(appendEmailLog(link.emailLog, entry)) });
    } catch (error) {
      log.warn?.('[checklist-email] could not record the email on the link', { rowId, error });
    }
  }

  const reason = (error) => String(error?.message ?? error ?? 'Not sent').split('\n')[0].slice(0, 300);

  /**
   * `link` as it stood before signing (its log is carried on), `values` what
   * was signed, `signatureBytes` the PNG or null.
   */
  async function onSigned({ rowId, link, values, signedOn, origin, signatureBytes }) {
    if (!from) return { skipped: 'no sender configured' };
    const { enabled, to } = await settings();
    if (!enabled) return { skipped: 'off' };

    const at = new Date(now()).toISOString();
    try {
      await send({ link, values, signedOn, origin, signatureBytes, resent: null, to });
      await record(rowId, link, { at, kind: EMAIL_KINDS.SIGNED, to, ok: true });
      return { ok: true, to };
    } catch (error) {
      log.error('[checklist-email] not sent on signing', { rowId, error });
      await record(rowId, link, { at, kind: EMAIL_KINDS.SIGNED, to, ok: false, error: reason(error) });
      return { ok: false, error: reason(error) };
    }
  }

  async function resend({ id, by, origin }) {
    if (!from) return { ok: false, status: 503, error: 'Emails are not set up on the server yet (no sending mailbox).' };
    const row = await graph.getLink(id);
    if (!row) return { ok: false, status: 404, error: 'That checklist is no longer on the list.' };
    const link = fromLinkItem({ ...row.fields, id: row.id });
    if (link.status !== LINK_STATUS.SIGNED || !link.submitted) {
      return { ok: false, status: 409, error: 'Only a signed checklist can be emailed.' };
    }

    const { to } = await settings();
    if (!to.length) {
      return { ok: false, status: 409, error: 'Nobody to send it to — add the group in Email settings first.' };
    }

    let signatureBytes = null;
    if (link.signatureFile) {
      try {
        signatureBytes = await graph.readSignature(link.signatureFile);
      } catch (error) {
        log.warn?.('[checklist-email] signature could not be read for a resend', error);
      }
    }

    const at = now();
    const iso = new Date(at).toISOString();
    try {
      await send({
        link, values: link.submitted, signedOn: link.signedOn, origin, signatureBytes, resent: { by, at }, to,
      });
    } catch (error) {
      log.error('[checklist-email] resend failed', { id, error });
      await record(row.id, link, { at: iso, kind: EMAIL_KINDS.RESENT, to, by, ok: false, error: reason(error) });
      return { ok: false, status: 502, error: `The email was not sent: ${reason(error)}` };
    }
    await record(row.id, link, { at: iso, kind: EMAIL_KINDS.RESENT, to, by, ok: true });
    return { ok: true, to };
  }

  async function test({ by, origin }) {
    if (!from) return { ok: false, status: 503, error: 'Emails are not set up on the server yet (no sending mailbox).' };
    const { to } = await settings();
    if (!to.length) return { ok: false, status: 409, error: 'Save at least one address first.' };
    const mail = buildTestEmail({ origin, to, by });
    try {
      await graph.sendMail({ from, to, subject: mail.subject, html: mail.html });
    } catch (error) {
      log.error('[checklist-email] test failed', error);
      return { ok: false, status: 502, error: `The test was not sent: ${reason(error)}` };
    }
    return { ok: true, to };
  }

  return { onSigned, resend, test };
}

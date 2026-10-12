import { describe, it, expect } from 'vitest';
import { createFakeGraph } from './fakeGraph.js';
import { createLinkApi } from './checklistLinkApi.js';
import { createChecklistMailer } from './checklistMailer.js';
import { buildChecklistEmail, listedItems } from './checklistEmail.js';
import { createCallerCheck } from './mailHandler.js';
import { toLinkItem, fromLinkItem, LINK_STATUS } from '../src/features/forms/links/linkSchema.js';
import { IN, OUT, INDIVIDUAL } from '../src/features/forms/checklistForm.js';
import {
  normaliseEmailSettings, addRecipient, MAX_RECIPIENTS,
} from '../src/features/forms/links/emailSettings.js';
import { appendEmailLog, emailStatus, MAX_LOG_ENTRIES } from '../src/features/forms/links/emailLog.js';

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
const NOW = Date.parse('2026-10-07T05:37:00Z');
const quiet = { warn() {}, error() {} };

const linkRow = (code, formMode, preset = {}, extra = {}) => ({
  ...toLinkItem({
    code, formMode, preset, editable: [], expiresOn: new Date(NOW + 86400000).toISOString(),
  }, { createdByName: 'Ashraf (IT)' }),
  ...extra,
});

const signing = (values = {}) => ({
  values: {
    employeeName: 'Chaw Voon Pin',
    employeeNo: 'SB005',
    position: 'Head of Finance',
    entity: 'PMW',
    department: 'Finance',
    formDate: '2026-10-07',
    checkedItems: ['Laptop', 'Mouse'],
    serialNumbers: '5CD1234XYZ',
    signature: PNG,
    ...values,
  },
});

function setup({ settings, fail, links } = {}) {
  const graph = createFakeGraph({
    links: links ?? [linkRow('CodeOut123', OUT)],
    settings: settings ?? { checklistEmail: { enabled: true, to: ['it-assets@pmw.test'] } },
    fail,
  });
  const mailer = createChecklistMailer({
    graph, from: 'portal@pmw.test', now: () => NOW, log: quiet,
  });
  const api = createLinkApi({
    graph, mailer, now: () => NOW, log: quiet,
  });
  return { graph, mailer, api };
}

const linkOf = (graph, id = 1) => fromLinkItem({ ...graph.rows.get(id).fields, id });

describe('email on signing', () => {
  it('sends one email to the saved group, with the signature inline, and logs it', async () => {
    const { graph, api } = setup();
    const answer = await api.submit('CodeOut123', signing(), { origin: 'https://it.pmw.test' });
    expect(answer.status).toBe(200);

    expect(graph.sent).toHaveLength(1);
    const [mail] = graph.sent;
    expect(mail.from).toBe('portal@pmw.test');
    expect(mail.to).toEqual(['it-assets@pmw.test']);
    expect(mail.subject).toBe('Signed: OUT checklist — Chaw Voon Pin (SB005)');
    expect(mail.attachments).toHaveLength(1);
    expect(mail.attachments[0].contentId).toBe('signature');
    expect(mail.html).toContain('cid:signature');
    expect(mail.html).toContain('https://it.pmw.test/c/CodeOut123');

    const { emailLog } = linkOf(graph);
    expect(emailLog).toEqual([{
      at: new Date(NOW).toISOString(), kind: 'signed', to: ['it-assets@pmw.test'], by: '', ok: true,
    }]);
  });

  it('sends nothing while emails are off', async () => {
    const { graph, api } = setup({ settings: { checklistEmail: { enabled: false, to: ['a@pmw.test'] } } });
    await api.submit('CodeOut123', signing());
    expect(graph.sent).toHaveLength(0);
    expect(linkOf(graph).emailLog).toEqual([]);
  });

  it('sends nothing when nobody ever saved the setting', async () => {
    const { graph, api } = setup({ settings: {} });
    await api.submit('CodeOut123', signing());
    expect(graph.sent).toHaveLength(0);
  });

  it('a failed send never fails the signing, and is logged as failed', async () => {
    const { graph, api } = setup({ fail: { sendMail: true } });
    const answer = await api.submit('CodeOut123', signing());
    expect(answer.status).toBe(200);
    expect(linkOf(graph).status).toBe(LINK_STATUS.SIGNED);
    const status = emailStatus(linkOf(graph));
    expect(status.state).toBe('failed');
    expect(status.last.error).toMatch(/403/);
  });

  it('a mailer that throws outright still leaves the checklist signed', async () => {
    const graph = createFakeGraph({ links: [linkRow('CodeOut123', OUT)] });
    const mailer = { onSigned: async () => { throw new Error('boom'); } };
    const api = createLinkApi({ graph, mailer, now: () => NOW, log: quiet });
    expect((await api.submit('CodeOut123', signing())).status).toBe(200);
    expect(graph.checklists).toHaveLength(1);
  });

  it('sends nothing for a submission that was refused', async () => {
    const { graph, api } = setup();
    const answer = await api.submit('CodeOut123', signing({ signature: null }));
    expect(answer.status).toBe(422);
    expect(graph.sent).toHaveLength(0);
  });
});

describe('resend', () => {
  it('emails an OLD signed checklist (signed while emails were off) to the saved group', async () => {
    const { graph, api, mailer } = setup({ settings: { checklistEmail: { enabled: false, to: ['a@pmw.test'] } } });
    await api.submit('CodeOut123', signing());
    expect(graph.sent).toHaveLength(0);

    const result = await mailer.resend({ id: 1, by: 'Ashraf', origin: 'https://it.pmw.test' });
    expect(result).toEqual({ ok: true, to: ['a@pmw.test'] });
    expect(graph.sent[0].subject).toMatch(/^\[Resent\] Signed: OUT checklist/);
    expect(graph.sent[0].html).toContain('Resent by Ashraf');
    expect(graph.sent[0].attachments).toHaveLength(1);
    expect(emailStatus(linkOf(graph)).last).toMatchObject({ kind: 'resent', by: 'Ashraf', ok: true });
  });

  it('refuses a link that is not signed', async () => {
    const { mailer } = setup();
    expect(await mailer.resend({ id: 1, by: 'X' })).toMatchObject({ ok: false, status: 409 });
  });

  it('refuses when no group is saved', async () => {
    const { api, mailer } = setup({ settings: {} });
    await api.submit('CodeOut123', signing());
    expect(await mailer.resend({ id: 1, by: 'X' })).toMatchObject({ ok: false, status: 409 });
  });

  it('answers 404 for a link that is gone', async () => {
    const { mailer } = setup();
    expect(await mailer.resend({ id: 99, by: 'X' })).toMatchObject({ ok: false, status: 404 });
  });

  it('says so when the server has no sending mailbox', async () => {
    const graph = createFakeGraph();
    const mailer = createChecklistMailer({ graph, from: '', log: quiet });
    expect(await mailer.resend({ id: 1 })).toMatchObject({ ok: false, status: 503 });
    expect(await mailer.onSigned({ rowId: 1, link: {}, values: {} })).toEqual({ skipped: 'no sender configured' });
  });
});

describe('the email itself', () => {
  const link = fromLinkItem({ ...linkRow('CodeAbc123', IN), id: 7 });

  it('escapes whatever the employee typed', () => {
    const { html, subject } = buildChecklistEmail({
      link,
      values: { employeeName: '<img src=x onerror=alert(1)>', otherRemarks: '<b>hi</b>', checkedItems: ['Laptop'] },
      signedOn: new Date(NOW).toISOString(),
      origin: 'https://it.pmw.test',
    });
    expect(html).not.toContain('<img src=x');
    expect(html).toContain('&lt;img src=x onerror=alert(1)&gt;');
    expect(html).not.toContain('<b>hi</b>');
    // A subject is plain text; nothing to escape there.
    expect(subject).toContain('<img src=x');
  });

  it('names the till and the handover rows on a till checklist', () => {
    const till = { ...link, handovers: { kind: 'issue', ids: [1, 2, 3] } };
    const { html } = buildChecklistEmail({ link: till, values: { checkedItems: [] }, signedOn: '2026-10-07T05:37:00Z' });
    expect(html).toContain('Handed out at the till by Ashraf (IT).');
    expect(html).toContain('3 handover records');
  });

  it('says when IT edited it after signing', () => {
    const edited = { ...link, editedBy: 'Ashraf', editedOn: '2026-10-08T01:00:00Z' };
    const { html } = buildChecklistEmail({ link: edited, values: {}, signedOn: '2026-10-07T05:37:00Z' });
    expect(html).toContain('Edited by Ashraf after signing');
  });

  it('lists a request with quantities, and IN/OUT as ticked items', () => {
    expect(listedItems(INDIVIDUAL, { items: [{ item: 'HDMI Cable', quantity: 2 }, { item: '', quantity: 1 }] }))
      .toEqual([{ name: 'HDMI Cable', quantity: 2 }]);
    expect(listedItems(OUT, { checkedItems: ['Laptop'] })).toEqual([{ name: 'Laptop', quantity: null }]);
  });

  it('shows the form date as the day picked, never shifted', () => {
    const { html } = buildChecklistEmail({ link, values: { formDate: '2026-10-07' }, signedOn: '2026-10-06T23:30:00Z' });
    expect(html).toContain('07/10/2026');
  });

  it('says plainly when no signature could be attached', () => {
    const { html } = buildChecklistEmail({ link, values: {}, signedOn: '2026-10-07T05:37:00Z', hasSignature: false });
    expect(html).not.toContain('cid:signature');
    expect(html).toContain('No signature image could be attached');
  });
});

describe('who may ask for a resend', () => {
  const tenantId = 'tenant-1';
  const portalClientId = 'portal-app';
  const token = (claims) => `x.${Buffer.from(JSON.stringify(claims)).toString('base64url')}.sig`;
  const good = {
    tid: tenantId, appid: portalClientId, aud: '00000003-0000-0000-c000-000000000000', exp: NOW / 1000 + 600,
  };
  const graphSays = (ok) => async () => ({
    ok, json: async () => ({ displayName: 'Ashraf', mail: 'ashraf@pmw.test' }),
  });
  const check = (ok = true) => createCallerCheck({
    tenantId, portalClientId, fetchImpl: graphSays(ok), now: () => NOW,
  });

  it('accepts a token Graph accepts, from this tenant, issued to the portal', async () => {
    expect(await check()(token(good))).toEqual({ name: 'Ashraf', email: 'ashraf@pmw.test' });
  });

  it('refuses a token Graph refuses', async () => {
    expect(await check(false)(token(good))).toBeNull();
  });

  it('refuses another tenant, another app, another audience, or an expired token', async () => {
    expect(await check()(token({ ...good, tid: 'other' }))).toBeNull();
    expect(await check()(token({ ...good, appid: 'graph-explorer' }))).toBeNull();
    expect(await check()(token({ ...good, aud: 'https://pmw.sharepoint.com' }))).toBeNull();
    expect(await check()(token({ ...good, exp: NOW / 1000 - 1 }))).toBeNull();
  });

  it('refuses nothing at all, or rubbish', async () => {
    expect(await check()('')).toBeNull();
    expect(await check()('not-a-token')).toBeNull();
  });
});

describe('settings and log rules', () => {
  it('reads a stored setting, and treats anything odd as off', () => {
    expect(normaliseEmailSettings('{"enabled":true,"to":["IT@PMW.test","it@pmw.test","nope"]}'))
      .toEqual({ enabled: true, to: ['it@pmw.test'] });
    expect(normaliseEmailSettings('{"enabled":true,"to":[]}')).toEqual({ enabled: false, to: [] });
    expect(normaliseEmailSettings('garbage')).toEqual({ enabled: false, to: [] });
    expect(normaliseEmailSettings(null)).toEqual({ enabled: false, to: [] });
  });

  it('checks an address before it goes on the list', () => {
    expect(addRecipient([], ' IT-Assets@PMW.test ')).toEqual({ address: 'it-assets@pmw.test' });
    expect(addRecipient(['a@b.co'], 'a@b.co').error).toMatch(/already/);
    expect(addRecipient([], 'not an address').error).toMatch(/not an email/);
    expect(addRecipient(Array.from({ length: MAX_RECIPIENTS }, (_, i) => `p${i}@b.co`), 'x@b.co').error).toMatch(/At most/);
  });

  it('keeps the log bounded, and the newest attempt decides the status', () => {
    let log = [];
    for (let i = 0; i < MAX_LOG_ENTRIES + 5; i += 1) {
      log = appendEmailLog(log, { at: new Date(NOW + i).toISOString(), kind: 'resent', ok: true });
    }
    expect(log).toHaveLength(MAX_LOG_ENTRIES);
    expect(emailStatus(log).state).toBe('sent');
    log = appendEmailLog(log, { at: new Date(NOW).toISOString(), kind: 'resent', ok: false, error: 'x' });
    expect(emailStatus(log).state).toBe('failed');
    expect(emailStatus(log).lastSent.ok).toBe(true);
    expect(emailStatus([]).state).toBe('none');
  });
});

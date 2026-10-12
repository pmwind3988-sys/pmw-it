import { FORM_MODES, INDIVIDUAL, IN, OUT } from '../src/features/forms/checklistForm.js';
import { optionLabel } from '../src/features/forms/formOptions.js';
import { formatMYT } from '../src/utils/malaysiaTime.js';

/**
 * The email sent when a shared checklist is signed, as `{ subject, html, text }`.
 *
 * Pure, so what the group receives is tested without a mailbox
 * (`checklistEmail.test.js`). Laid out in tables with inline styles, because
 * that is what Outlook renders; anything cleverer falls apart there.
 *
 * EVERY value the employee typed is escaped. The public page is anonymous, so
 * a name field holding markup must arrive as text, not as part of the email.
 *
 * The signature is NOT linked: it lives in SharePoint, which no mail client
 * can read, and that is why the old notification showed a broken image. The
 * sender attaches it inline and this refers to it as `cid:signature`.
 */

export const SIGNATURE_CID = 'signature';

const BRAND = '#0078d4';
const BRAND_DEEP = '#004e86';
const INK = '#101828';
const SOFT = '#566479';
const LINE = '#e3e9f0';
const FONT = "'Segoe UI', system-ui, -apple-system, Arial, sans-serif";

const escapeHtml = (value) => String(value ?? '')
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  .replace(/'/g, '&#39;');

const clean = (value) => String(value ?? '').trim();

const MODE_WORDS = {
  [IN]: { subject: 'IN checklist', heading: 'items received', listTitle: 'Received' },
  [OUT]: { subject: 'OUT checklist', heading: 'items handed back', listTitle: 'Handed back' },
  [INDIVIDUAL]: { subject: 'Individual request', heading: 'request received', listTitle: 'Requested' },
};

const modeOf = (formMode) => FORM_MODES.find((mode) => mode.value === formMode) ?? null;

const longDate = (at) => {
  const time = typeof at === 'number' ? at : Date.parse(at);
  if (!Number.isFinite(time)) return '';
  const day = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kuala_Lumpur', weekday: 'long', day: 'numeric', month: 'long', year: 'numeric',
  }).format(new Date(time));
  return `${day} at ${formatMYT(time, 'time12')}`;
};

const shortDateTime = (at) => {
  const time = typeof at === 'number' ? at : Date.parse(at);
  return Number.isFinite(time) ? formatMYT(time, 'datetime12') : '';
};

// `yyyy-mm-dd` as the employee picked it; a day, so never shifted by a zone.
const formDay = (value) => {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(clean(value));
  return match ? `${match[3]}/${match[2]}/${match[1]}` : clean(value);
};

const linesOf = (value) => clean(value).split(/\r?\n/).map((line) => line.trim()).filter(Boolean);

/** What the checklist lists, as `[{ name, quantity }]`: ticked for IN/OUT, typed for a request. */
export function listedItems(formMode, values = {}) {
  if (formMode === INDIVIDUAL) {
    return (Array.isArray(values.items) ? values.items : [])
      .map((row) => ({ name: clean(row?.item), quantity: Number(row?.quantity) || 1 }))
      .filter((row) => row.name);
  }
  return (Array.isArray(values.checkedItems) ? values.checkedItems : [])
    .map((name) => ({ name: clean(name), quantity: null }))
    .filter((row) => row.name);
}

const tillNote = (link) => {
  const kind = link?.handovers?.kind;
  if (!kind) return '';
  const by = clean(link.createdByName) || 'IT';
  const count = Array.isArray(link.handovers.ids) ? link.handovers.ids.length : 0;
  const what = kind === 'return' ? `Taken back at the till by ${by}.` : `Handed out at the till by ${by}.`;
  if (!count) return what;
  return `${what} The signature has also been added to the ${count} ${kind === 'return' ? 'return' : 'handover'} record${count === 1 ? '' : 's'}.`;
};

const editNote = (link) => {
  if (!clean(link?.editedBy)) return '';
  const on = link.editedOn ? ` on ${shortDateTime(link.editedOn)}` : '';
  return `Edited by ${clean(link.editedBy)} after signing${on}. The values below are the corrected ones; the signed copy notes the change.`;
};

/**
 * @param {object} p
 * @param {object} p.link       the link (fromLinkItem shape)
 * @param {object} p.values     what was signed (as stored on the link)
 * @param {string} p.signedOn   ISO instant
 * @param {string} p.origin     the portal's address, for the buttons
 * @param {boolean} p.hasSignature  whether a signature is attached inline
 * @param {{ by: string, at: number|string }|null} p.resent
 * @param {string[]} p.to       who it goes to, named in the footer
 */
export function buildChecklistEmail({
  link, values = {}, signedOn, origin, hasSignature = false, resent = null, to = [],
}) {
  const formMode = link?.formMode ?? values.formMode;
  const mode = modeOf(formMode);
  const words = MODE_WORDS[formMode] ?? { subject: 'Checklist', heading: 'checklist signed', listTitle: 'Items' };
  const options = link?.options ?? null;

  const name = clean(values.employeeName) || clean(link?.employeeName) || 'An employee';
  const number = clean(values.employeeNo);
  const entity = clean(optionLabel(options, values, 'entity'));
  const department = clean(optionLabel(options, values, 'department'));
  const items = listedItems(formMode, values);
  const serials = linesOf(values.serialNumbers);
  const remarks = clean(values.otherRemarks);
  const base = String(origin ?? '').replace(/\/$/, '');
  const copyUrl = `${base}/c/${encodeURIComponent(link?.code ?? '')}`;
  const printUrl = `${copyUrl}?print=1`;
  const listUrl = `${base}/asset-checklist/links?open=${encodeURIComponent(link?.id ?? '')}`;
  const isRequest = formMode === INDIVIDUAL;

  const subject = `${resent ? '[Resent] ' : ''}Signed: ${words.subject} — ${name}${number ? ` (${number})` : ''}`;

  const notes = [
    resent ? { tone: 'amber', text: `Resent by ${clean(resent.by) || 'IT'} on ${shortDateTime(resent.at)}. First sent when it was signed.` } : null,
    tillNote(link) ? { tone: 'blue', text: tillNote(link) } : null,
    editNote(link) ? { tone: 'amber', text: editNote(link) } : null,
  ].filter(Boolean);

  const facts = [
    ['Name', name],
    ['Employee no.', number || '—'],
    ['Position', clean(values.position) || '—'],
    ['Entity · Department', [entity, department].filter(Boolean).join(' · ') || '—'],
    ['Date on the form', formDay(values.formDate) || '—'],
    ['Signed on', shortDateTime(signedOn) || '—'],
  ];

  // ---- HTML -------------------------------------------------------------
  const e = escapeHtml;
  const label = (text) => `<div style="font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:${SOFT};padding:0 0 8px">${e(text)}</div>`;
  const section = (inner, pad = '12px 32px') => `<tr><td style="padding:${pad}">${inner}</td></tr>`;

  const noteBox = ({ tone, text }) => {
    const [bg, border, ink] = tone === 'amber' ? ['#fffbeb', '#f3dca0', '#7a4a00'] : ['#eaf4fd', '#cfe3f7', BRAND_DEEP];
    return `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:12px 0 0"><tr><td style="background:${bg};border:1px solid ${border};border-radius:8px;padding:12px 14px;font-size:13px;line-height:1.5;color:${ink}">${e(text)}</td></tr></table>`;
  };

  const factCells = facts.map(([key, value]) => `<td width="50%" valign="top" style="padding:8px 12px"><div style="font-size:12px;color:${SOFT}">${e(key)}</div><div style="font-size:15px;font-weight:600;color:${INK}">${e(value)}</div></td>`);
  const factRows = [];
  for (let i = 0; i < factCells.length; i += 2) factRows.push(`<tr>${factCells[i]}${factCells[i + 1] ?? '<td></td>'}</tr>`);

  const itemRows = items.length
    ? items.map((item, index) => `<tr><td style="padding:12px 16px;${index ? `border-top:1px solid ${LINE};` : ''}font-size:15px;font-weight:600;color:${INK}">${isRequest ? '' : `<span style="color:#12a150">&#10003;</span>&nbsp; `}${e(item.name)}</td>${isRequest ? `<td align="right" style="padding:12px 16px;${index ? `border-top:1px solid ${LINE};` : ''}font-size:15px;color:${INK}">${e(item.quantity)}</td>` : ''}</tr>`).join('')
    : `<tr><td style="padding:12px 16px;font-size:14px;color:${SOFT}">Nothing listed</td></tr>`;
  const itemHead = isRequest
    ? `<tr><td style="padding:10px 16px;background:#f4f7fb;border-bottom:1px solid ${LINE};font-size:12px;font-weight:600;color:${SOFT}">Item</td><td align="right" style="padding:10px 16px;background:#f4f7fb;border-bottom:1px solid ${LINE};font-size:12px;font-weight:600;color:${SOFT}">Qty</td></tr>`
    : '';

  const signatureCell = hasSignature
    ? `<img src="cid:${SIGNATURE_CID}" alt="Signature of ${e(name)}" width="220" style="display:block;max-width:220px;height:auto;border:0">`
    : `<span style="font-size:13px;color:${SOFT}">No signature image could be attached. Open the signed copy to see it.</span>`;

  const button = (href, text, primary) => (primary
    ? `<a href="${e(href)}" style="display:inline-block;background:${BRAND};color:#ffffff;text-decoration:none;font-size:14px;font-weight:600;padding:12px 20px;border-radius:6px;margin:0 8px 8px 0">${e(text)}</a>`
    : `<a href="${e(href)}" style="display:inline-block;background:#ffffff;color:${BRAND_DEEP};border:1px solid #b9cde4;text-decoration:none;font-size:14px;font-weight:600;padding:11px 20px;border-radius:6px;margin:0 8px 8px 0">${e(text)}</a>`);

  const recipients = to.length ? to.join(', ') : 'the checklist email group';
  const footer = resent
    ? `Resent on request from the IT Service Portal to ${recipients}.`
    : `Sent once, when the checklist was signed, to ${recipients}. The list is changed in the IT Service Portal under Shared checklists › Email settings.`;

  const html = `<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${e(subject)}</title></head>
<body style="margin:0;padding:0;background:#eef2f7;font-family:${FONT};color:${INK}">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eef2f7"><tr><td align="center" style="padding:24px 12px">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;background:#ffffff;border:1px solid #dfe5ee;border-radius:10px;font-family:${FONT}">
<tr><td style="background:${BRAND};color:#ffffff;padding:26px 32px;border-radius:10px 10px 0 0">
<div style="font-size:12px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:#e6f0fa">PMW IT · Asset checklist</div>
<div style="font-size:26px;font-weight:700;line-height:1.2;padding:8px 0 10px;color:#ffffff">Checklist signed — ${e(words.heading)}</div>
<span style="display:inline-block;background:#ffffff;color:${BRAND_DEEP};font-size:12px;font-weight:700;padding:4px 10px;border-radius:999px">${e(mode?.label ?? formMode ?? '')}</span>
<span style="font-size:13px;color:#e6f0fa">&nbsp; ${e(mode?.description ?? '')}</span>
</td></tr>
${section(`<div style="font-size:15px;line-height:1.55;color:${INK}"><b>${e(name)}</b> signed ${isRequest ? 'an individual request' : `an ${e(mode?.label ?? '')} checklist`} on <b>${e(longDate(signedOn))}</b>.</div>${notes.map(noteBox).join('')}`, '26px 32px 8px')}
${section(`${label('Employee')}<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid ${LINE};border-radius:8px">${factRows.join('')}</table>`, '16px 32px 12px')}
${section(`${label(`${words.listTitle} (${items.length}${isRequest ? ` line${items.length === 1 ? '' : 's'}` : ''})`)}<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid ${LINE};border-radius:8px">${itemHead}${itemRows}</table>`)}
${section(`${label('Serial numbers')}${serials.length ? `<div style="font-family:Consolas,'Cascadia Mono',monospace;font-size:13px;line-height:1.7;color:#344054">${serials.map(e).join('<br>')}</div>` : `<div style="font-size:14px;color:${SOFT}">None</div>`}`)}
${section(`${label('Other remarks')}<div style="font-size:14px;line-height:1.5;color:${remarks ? INK : SOFT}">${remarks ? linesOf(remarks).map(e).join('<br>') : 'None'}</div>`)}
${section(`${label('Employee signature')}<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid ${LINE};border-radius:8px;background:#fbfcfe"><tr><td style="padding:14px 18px" valign="bottom">${signatureCell}</td><td align="right" valign="bottom" style="padding:14px 18px;font-size:12px;line-height:1.5;color:${SOFT}">${e(name)}<br>${e(shortDateTime(signedOn))}</td></tr></table>`, '12px 32px 8px')}
${section(`${button(copyUrl, 'Open signed copy', true)}${button(printUrl, 'Print / Save as PDF', false)}<a href="${e(listUrl)}" style="display:inline-block;color:${BRAND};font-size:14px;font-weight:600;padding:12px 4px;text-decoration:none">See all checklists</a>`, '22px 32px 20px')}
<tr><td style="background:#f4f7fb;border-top:1px solid ${LINE};padding:16px 32px;font-size:12px;line-height:1.6;color:${SOFT};border-radius:0 0 10px 10px">${e(footer)}<br>&copy; ${new Date(Date.parse(signedOn) || Date.now()).getUTCFullYear()} PMW Group</td></tr>
</table>
</td></tr></table>
</body></html>`;

  // ---- Plain text, for clients that show no HTML -------------------------
  const text = [
    `Checklist signed — ${words.heading} (${mode?.label ?? formMode ?? ''})`,
    '',
    `${name} signed ${isRequest ? 'an individual request' : `an ${mode?.label ?? ''} checklist`} on ${longDate(signedOn)}.`,
    ...notes.map((note) => note.text),
    '',
    ...facts.map(([key, value]) => `${key}: ${value}`),
    '',
    `${words.listTitle}:`,
    ...(items.length ? items.map((item) => `  - ${item.name}${isRequest ? ` × ${item.quantity}` : ''}`) : ['  (nothing listed)']),
    '',
    'Serial numbers:',
    ...(serials.length ? serials.map((line) => `  ${line}`) : ['  None']),
    '',
    `Other remarks: ${remarks || 'None'}`,
    '',
    `Signed copy: ${copyUrl}`,
    `All checklists: ${listUrl}`,
    '',
    footer,
  ].join('\n');

  return { subject, html, text };
}

/** The test email: says what it is, and who would be emailed. */
export function buildTestEmail({ origin, to = [], by = '' }) {
  const base = String(origin ?? '').replace(/\/$/, '');
  const subject = 'Test: checklist signing emails';
  const lines = [
    `${clean(by) || 'Somebody in IT'} sent this test from the IT Service Portal.`,
    `When an employee signs a shared checklist, an email like the real one goes to: ${to.join(', ')}.`,
    'Nothing was signed; there is nothing to do.',
  ];
  const html = `<!doctype html><html lang="en"><head><meta charset="utf-8"><title>${escapeHtml(subject)}</title></head>
<body style="margin:0;padding:24px 12px;background:#eef2f7;font-family:${FONT};color:${INK}">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" align="center" style="width:100%;max-width:600px;background:#ffffff;border:1px solid #dfe5ee;border-radius:10px">
<tr><td style="background:${BRAND};color:#ffffff;padding:22px 32px;border-radius:10px 10px 0 0;font-size:22px;font-weight:700">Checklist emails are working</td></tr>
<tr><td style="padding:22px 32px;font-size:15px;line-height:1.6">${lines.map(escapeHtml).join('<br><br>')}</td></tr>
<tr><td style="padding:0 32px 24px"><a href="${escapeHtml(`${base}/asset-checklist/links`)}" style="color:${BRAND};font-weight:600;text-decoration:none">Open Shared checklists</a></td></tr>
</table></body></html>`;
  return { subject, html, text: lines.join('\n\n') };
}

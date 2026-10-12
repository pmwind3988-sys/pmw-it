/**
 * An in-memory stand-in for `createGraph`, with the same methods.
 *
 * Used by the handler tests, and by `npm run dev` when `CHECKLIST_FAKE=1` so
 * the public page can be driven end to end on a machine with no app secret.
 * Its eTags change on every write, which is what lets the tests prove that
 * two submissions cannot both claim one link.
 */

import { ConflictError } from './errors.js';

export { ConflictError };

export function createFakeGraph({
  links = [], handovers = {}, settings = {}, fail = {},
} = {}) {
  let version = 1;
  let nextChecklistId = 100;
  const rows = new Map();
  const checklists = [];
  const files = new Map();
  // Every email "sent": { from, to, subject, html, attachments }.
  const sent = [];
  const settingRows = new Map(Object.entries(settings));
  // Till handover rows by id: { IssueSignature, ReturnSignature, ... }.
  const handoverRows = new Map(Object.entries(handovers).map(([id, fields]) => [String(id), { ...fields }]));

  const store = (id, fields) => {
    version += 1;
    rows.set(id, { id: String(id), eTag: `"v${version}"`, fields: { ...fields, Modified: new Date().toISOString() } });
  };

  links.forEach((fields, index) => store(index + 1, fields));

  return {
    rows,
    checklists,
    files,
    handoverRows,
    sent,
    settingRows,

    async findLink(code) {
      const row = [...rows.values()].find((entry) => entry.fields.Title === code);
      return row ? structuredClone(row) : null;
    },

    async updateLink(id, fields, { eTag } = {}) {
      if (fail.updateLink) throw new Error('updateLink failed');
      const row = rows.get(Number(id));
      if (!row) throw new Error('No such link');
      if (eTag && eTag !== row.eTag) throw new ConflictError();
      store(Number(id), { ...row.fields, ...fields });
    },

    async createChecklist(fields) {
      if (fail.createChecklist) throw new Error('createChecklist failed');
      nextChecklistId += 1;
      checklists.push({ id: nextChecklistId, fields });
      return { id: nextChecklistId };
    },

    async updateChecklist(id, fields) {
      if (fail.updateChecklist) throw new Error('updateChecklist failed');
      const row = checklists.find((entry) => entry.id === Number(id));
      if (!row) throw new Error('No such checklist row');
      row.fields = { ...row.fields, ...fields };
    },

    async uploadSignature(fileName, bytes) {
      if (fail.uploadSignature) throw new Error('uploadSignature failed');
      files.set(fileName, bytes);
      return { serverRelativeUrl: `/sites/IThelpdesk/Signatures/${fileName}` };
    },

    async signHandover(id, field, url) {
      if (fail.signHandover) throw new Error('signHandover failed');
      const row = handoverRows.get(String(id));
      if (!row) throw new Error('No such handover row');
      if (String(row[field] ?? '').trim()) return false;
      row[field] = url;
      return true;
    },

    async getLink(id) {
      const row = rows.get(Number(id));
      return row ? structuredClone(row) : null;
    },

    async readSetting(key) {
      const value = settingRows.get(key);
      if (value === undefined) return null;
      return typeof value === 'string' ? value : JSON.stringify(value);
    },

    async sendMail(mail) {
      if (fail.sendMail) throw new Error('sendMail failed (403): ErrorAccessDenied');
      sent.push(structuredClone(mail));
    },

    async readSignature(fileName) {
      return files.get(fileName) ?? null;
    },
  };
}

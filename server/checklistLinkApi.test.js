import { describe, it, expect, vi } from 'vitest';
import { createLinkApi } from './checklistLinkApi.js';
import { createFakeGraph, ConflictError } from './fakeGraph.js';
import { toLinkItem, fromLinkItem, LINK_STATUS } from '../src/features/forms/links/linkSchema.js';
import { reopenFields } from '../src/features/forms/links/linkChanges.js';
import { IN } from '../src/features/forms/checklistForm.js';

const NOW = Date.parse('2026-10-05T04:00:00Z');
const DAY = 86400000;
const CODE = 'Ab3dE5gH9k';
const signature = 'data:image/png;base64,iVBORw0KGgo=';

const linkFields = (overrides = {}) => ({
  ...toLinkItem({
    code: CODE,
    formMode: IN,
    preset: {
      employeeName: 'Amir Hakim',
      employeeNo: 'E-1042',
      position: '',
      entity: 'PMW',
      department: 'Engineering',
      formDate: '2026-10-05',
      checkedItems: ['Laptop'],
      serialNumbers: 'SN-1',
    },
    editable: [],
    expiresOn: new Date(NOW + 3 * DAY).toISOString(),
  }, { createdByName: 'IT Desk', createdByEmail: 'it@pmw.test' }),
  ...overrides,
});

const setup = (overrides, fail) => {
  const graph = createFakeGraph({ links: [linkFields(overrides)], fail });
  const log = { error: vi.fn(), warn: vi.fn() };
  const api = createLinkApi({ graph, now: () => NOW, log });
  return { graph, api, log };
};

const answer = (overrides = {}) => ({
  values: { position: 'Engineer', employeeName: 'Somebody Else', signature, ...overrides },
});

describe('opening a link', () => {
  it('returns the pre-filled form and what may be changed', async () => {
    const { api } = setup();
    const { status, body } = await api.get(CODE);

    expect(status).toBe(200);
    expect(body.state).toBe('open');
    expect(body.formMode).toBe(IN);
    expect(body.values.employeeName).toBe('Amir Hakim');
    expect(body.editable).toEqual(['position', 'otherRemarks']);
  });

  it('gives an unknown, expired and cancelled link the same answer', async () => {
    const unknown = await setup().api.get('Zz9Zz9Zz9Z');
    const expired = await setup({ ExpiresOn: new Date(NOW - 1).toISOString() }).api.get(CODE);
    const cancelled = await setup({ LinkStatus: LINK_STATUS.CANCELLED }).api.get(CODE);
    const malformed = await setup().api.get("x' or 1");

    for (const result of [unknown, expired, cancelled, malformed]) {
      expect(result).toEqual({ status: 404, body: { state: 'gone' } });
    }
  });

  it('does not show who created the link or its internals', async () => {
    const { body } = await setup().api.get(CODE);
    const flat = JSON.stringify(body);
    expect(flat).not.toContain('it@pmw.test');
    expect(flat).not.toContain('IT Desk');
  });
});

describe('submitting a link', () => {
  it('writes the checklist row, keeps the locked values, and locks the link', async () => {
    const { api, graph } = setup();
    const { status, body } = await api.submit(CODE, answer());

    expect(status).toBe(200);
    expect(body.state).toBe('signed');

    expect(graph.checklists).toHaveLength(1);
    const row = graph.checklists[0].fields;
    expect(row.EmployeeName).toBe('Amir Hakim');
    expect(row.Position).toBe('Engineer');
    expect(row.FormMode).toBe(IN);
    expect(row.AssetMatrix).toBe('Laptop');
    expect(row.Department).toBe('Engineering');
    expect(row.SignatureUrl).toMatch(/^\/sites\/IThelpdesk\/Signatures\//);

    const link = graph.rows.get(1).fields;
    expect(link.LinkStatus).toBe(LINK_STATUS.SIGNED);
    expect(link.ChecklistId).toBe(graph.checklists[0].id);
    expect(JSON.parse(link.Submitted).position).toBe('Engineer');
    expect(JSON.parse(link.Submitted).signature).toBeUndefined();
  });

  it('then shows the signed copy, with its signature, on the same link', async () => {
    const { api } = setup();
    await api.submit(CODE, answer());
    const { status, body } = await api.get(CODE);

    expect(status).toBe(200);
    expect(body.state).toBe('signed');
    expect(body.values.position).toBe('Engineer');
    expect(body.signature).toBe(signature);
    expect(body.signedOn).toBe(new Date(NOW).toISOString());
  });

  it('refuses a second submission', async () => {
    const { api, graph } = setup();
    await api.submit(CODE, answer());
    const second = await api.submit(CODE, answer({ position: 'Changed' }));

    expect(second.status).toBe(409);
    expect(graph.checklists).toHaveLength(1);
  });

  it('refuses when another submission claimed the link first', async () => {
    const { api, graph } = setup();
    const realUpdate = graph.updateLink;
    graph.updateLink = vi.fn(async () => { throw new ConflictError(); });

    const { status } = await api.submit(CODE, answer());
    expect(status).toBe(409);
    expect(graph.checklists).toHaveLength(0);
    graph.updateLink = realUpdate;
  });

  it('says what is missing and saves nothing', async () => {
    const { api, graph } = setup();
    const { status, body } = await api.submit(CODE, { values: { position: 'Engineer' } });

    expect(status).toBe(422);
    expect(body.errors.signature).toBeTruthy();
    expect(graph.checklists).toHaveLength(0);
    expect(graph.rows.get(1).fields.LinkStatus).toBe(LINK_STATUS.WAITING);
  });

  it('puts the link back to waiting when the save fails, so it can be tried again', async () => {
    const { api, graph, log } = setup({}, { createChecklist: true });
    const { status, body } = await api.submit(CODE, answer());

    expect(status).toBe(502);
    expect(body.error).toMatch(/try again/i);
    expect(graph.rows.get(1).fields.LinkStatus).toBe(LINK_STATUS.WAITING);
    expect(log.error).toHaveBeenCalled();
  });

  it('refuses an expired link', async () => {
    const { api } = setup({ ExpiresOn: new Date(NOW - 1).toISOString() });
    expect((await api.submit(CODE, answer())).status).toBe(404);
  });
});

describe("a link that carries HR's entity and department lists", () => {
  const options = {
    entities: [{ value: 'PMWL', label: 'PMW Lighting' }],
    departments: { PMWL: [{ value: 'ENG', label: 'Engineering' }] },
  };

  it('hands the choices to the page and checks the answer against them', async () => {
    const { api, graph } = setup({
      Options: JSON.stringify(options),
      Preset: JSON.stringify({
        employeeName: 'Amir Hakim', employeeNo: 'E-1042', entity: '', formDate: '2026-10-05', checkedItems: ['Laptop'],
      }),
    });

    const opened = await api.get(CODE);
    expect(opened.body.options).toEqual(options);
    expect(opened.body.editable).toEqual(expect.arrayContaining(['entity', 'department']));

    const refused = await api.submit(CODE, answer({ entity: 'PMWL', department: 'Made Up' }));
    expect(refused.status).toBe(422);
    expect(refused.body.errors.department).toBeTruthy();

    const { status } = await api.submit(CODE, answer({ entity: 'PMWL', department: 'ENG' }));
    expect(status).toBe(200);
    expect(graph.checklists[0].fields).toMatchObject({ Entity: 'PMWL', Department: 'ENG' });
  });
});

describe('a reopened link', () => {
  const reopened = async ({ fail } = {}) => {
    const ctx = setup({}, fail);
    await ctx.api.submit(CODE, answer());
    const first = ctx.graph.checklists[0];
    const fields = ctx.graph.rows.get(1).fields;
    // What the portal's Reopen writes (`reopenFields`).
    const link = fromLinkItem({ ...fields, id: 1 });
    await ctx.graph.updateLink(1, reopenFields(link, NOW + 14 * DAY));
    return { ...ctx, first, fields };
  };

  it('opens with what was signed, and the signature already on it', async () => {
    const { api } = await reopened();
    const { body } = await api.get(CODE);
    expect(body.state).toBe('open');
    expect(body.values.position).toBe('Engineer');
    expect(body.existingSignature).toBe(signature);
  });

  it('updates the same record rather than adding a second one', async () => {
    const { api, graph, first } = await reopened();
    const { status } = await api.submit(CODE, answer({ position: 'Senior Engineer' }));

    expect(status).toBe(200);
    expect(graph.checklists).toHaveLength(1);
    expect(graph.checklists[0].id).toBe(first.id);
    expect(graph.checklists[0].fields.Position).toBe('Senior Engineer');
    expect(graph.rows.get(1).fields.ChecklistId).toBe(first.id);
  });

  it('keeps the signature it has when the employee does not sign again', async () => {
    const { api, graph, fields, first } = await reopened();
    const before = graph.files.size;
    const { status, body } = await api.submit(CODE, {
      values: { position: 'Senior Engineer' },
      keepSignature: true,
    });

    expect(status).toBe(200);
    expect(graph.files.size).toBe(before);
    expect(graph.checklists[0].fields.SignatureUrl).toBe(first.fields.SignatureUrl);
    expect(graph.checklists[0].fields.Position).toBe('Senior Engineer');
    expect(graph.rows.get(1).fields.SignatureFile).toBe(fields.SignatureFile);
    expect(body.signature).toBe(signature);
  });

  it('stores a new signature beside the old one when they sign again', async () => {
    const { api, graph } = await reopened();
    const before = graph.files.size;
    const again = 'data:image/png;base64,iVBORw0KGgoAAAAN';
    await api.submit(CODE, answer({ signature: again }));
    expect(graph.files.size).toBe(before + 1);
  });

  it('clears IT\'s "edited after signing" note, because the new signature covers it', async () => {
    const { api, graph, fields } = await reopened();
    await graph.updateLink(1, { EditedBy: 'IT Desk', EditedOn: new Date(NOW).toISOString() });
    const { status } = await api.submit(CODE, { values: JSON.parse(fields.Submitted), keepSignature: true });

    expect(status).toBe(200);

    expect(graph.checklists[0].fields.EditedAfterSigning).toBe('');
    expect(graph.rows.get(1).fields.EditedBy).toBe('');
  });

  it('cannot keep a signature on a link that never had one', async () => {
    const { api } = setup();
    const { status, body } = await api.submit(CODE, { values: { position: 'Engineer' }, keepSignature: true });
    expect(status).toBe(422);
    expect(body.errors.signature).toBeTruthy();
  });
});

describe('a checklist IT edited after signing', () => {
  it('says so on the signed copy', async () => {
    const { api, graph } = setup();
    await api.submit(CODE, answer());
    await graph.updateLink(1, { EditedBy: 'IT Desk', EditedOn: '2026-10-05T05:00:00.000Z' });

    const { body } = await api.get(CODE);
    expect(body.edited).toEqual({ by: 'IT Desk', on: '2026-10-05T05:00:00.000Z' });
  });

  it('says nothing when nobody edited it', async () => {
    const { api } = setup();
    await api.submit(CODE, answer());
    expect((await api.get(CODE)).body.edited).toBeNull();
  });
});

describe('a checklist made at the till', () => {
  const tillSetup = (handovers, rows, fail) => {
    const graph = createFakeGraph({ links: [linkFields({ Handovers: JSON.stringify(handovers) })], handovers: rows, fail });
    const log = { error: vi.fn(), warn: vi.fn() };
    return { graph, api: createLinkApi({ graph, now: () => NOW, log }), log };
  };

  it('puts the signature on the handover rows it was made with', async () => {
    const { graph, api } = tillSetup({ kind: 'issue', ids: [7, 8] }, { 7: {}, 8: {} });
    const { status } = await api.submit(CODE, answer());
    expect(status).toBe(200);
    const url = graph.handoverRows.get('7').IssueSignature;
    expect(url).toMatch(/^\/sites\/IThelpdesk\/Signatures\//);
    expect(graph.handoverRows.get('8').IssueSignature).toBe(url);
  });

  it('signs a return as "signed for (in)"', async () => {
    const { graph, api } = tillSetup({ kind: 'return', ids: [9] }, { 9: { IssueSignature: '/old.png' } });
    await api.submit(CODE, answer());
    expect(graph.handoverRows.get('9').ReturnSignature).toMatch(/Signatures/);
    expect(graph.handoverRows.get('9').IssueSignature).toBe('/old.png');
  });

  it('never writes over a signature already there', async () => {
    const { graph, api } = tillSetup({ kind: 'issue', ids: [7] }, { 7: { IssueSignature: '/signed-on-phone.png' } });
    await api.submit(CODE, answer());
    expect(graph.handoverRows.get('7').IssueSignature).toBe('/signed-on-phone.png');
  });

  it('still records the signed checklist when a handover row cannot be signed', async () => {
    const { graph, api, log } = tillSetup({ kind: 'issue', ids: [7] }, { 7: {} }, { signHandover: true });
    const { status } = await api.submit(CODE, answer());
    expect(status).toBe(200);
    expect(graph.checklists).toHaveLength(1);
    expect(log.warn).toHaveBeenCalled();
  });

  it('leaves an ordinary link alone', async () => {
    const { graph, api } = setup();
    await api.submit(CODE, answer());
    expect(graph.handoverRows.size).toBe(0);
  });
});

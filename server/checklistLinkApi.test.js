import { describe, it, expect, vi } from 'vitest';
import { createLinkApi } from './checklistLinkApi.js';
import { createFakeGraph, ConflictError } from './fakeGraph.js';
import { toLinkItem, LINK_STATUS } from '../src/features/forms/links/linkSchema.js';
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

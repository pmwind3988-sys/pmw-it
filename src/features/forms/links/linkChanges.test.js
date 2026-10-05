import { describe, it, expect } from 'vitest';
import { IN, INDIVIDUAL } from '../checklistForm.js';
import { LINK_STATUS, fromLinkItem, toLinkItem } from './linkSchema.js';
import {
  linkActions, planEdit, reopenFields, expiryFields, editNote, isReopened, endOfDay, dayOf,
} from './linkChanges.js';

const NOW = Date.parse('2026-10-05T04:00:00Z');
const DAY = 86400000;

const signedValues = {
  employeeName: 'Amir Hakim',
  employeeNo: 'E-1042',
  position: 'Engineer',
  entity: 'PMW',
  department: 'Engineering',
  formDate: '2026-10-05',
  checkedItems: ['Laptop'],
  items: [],
  serialNumbers: 'SN-1',
  otherRemarks: '',
};

const link = (overrides = {}) => ({
  id: 7,
  code: 'Ab3dE5gH9k',
  formMode: IN,
  preset: { employeeName: 'Amir Hakim' },
  editable: [],
  options: null,
  expiresOn: new Date(NOW + 3 * DAY).toISOString(),
  status: LINK_STATUS.WAITING,
  modified: new Date(NOW - DAY).toISOString(),
  submitted: null,
  checklistId: null,
  signatureFile: '',
  signedOn: '',
  editedBy: '',
  editedOn: '',
  originalSubmitted: null,
  ...overrides,
});

const signed = (overrides = {}) => link({
  status: LINK_STATUS.SIGNED,
  submitted: signedValues,
  checklistId: 41,
  signatureFile: 'In-sig.png',
  signedOn: new Date(NOW - DAY).toISOString(),
  ...overrides,
});

describe('what IT can do with a link', () => {
  it('offers expiry, cancel and delete while it waits', () => {
    expect(linkActions(link(), NOW)).toEqual(['copy', 'open', 'expiry', 'expireNow', 'cancel', 'delete']);
  });

  it('offers edit, reopen and delete once signed', () => {
    expect(linkActions(signed(), NOW)).toEqual(['view', 'edit', 'reopen', 'delete']);
  });

  it('offers reopen and delete once expired or cancelled', () => {
    const expired = link({ expiresOn: new Date(NOW - 1).toISOString() });
    expect(linkActions(expired, NOW)).toEqual(['reopen', 'delete']);
    expect(linkActions(link({ status: LINK_STATUS.CANCELLED }), NOW)).toEqual(['reopen', 'delete']);
  });

  it('leaves a link alone while somebody is in the middle of signing it', () => {
    const busy = link({ status: LINK_STATUS.SIGNING, modified: new Date(NOW - 1000).toISOString() });
    expect(linkActions(busy, NOW)).toEqual(['open']);
  });
});

describe('changing when a link expires', () => {
  it('sets the new date', () => {
    expect(expiryFields(NOW + 7 * DAY)).toEqual({ ExpiresOn: new Date(NOW + 7 * DAY).toISOString() });
  });

  it('runs through the whole of the day picked', () => {
    const end = new Date(endOfDay('2026-10-19'));
    expect([end.getFullYear(), end.getMonth(), end.getDate(), end.getHours(), end.getMinutes()])
      .toEqual([2026, 9, 19, 23, 59]);
    expect(dayOf(end.getTime())).toBe('2026-10-19');
    expect(Number.isNaN(endOfDay('next week'))).toBe(true);
  });

  it('refuses a date that is not a date', () => {
    expect(() => expiryFields('soon')).toThrow(/date/i);
  });
});

describe('reopening a link', () => {
  it('puts a signed link back to waiting, pre-filled with what was signed', () => {
    const fields = reopenFields(signed(), NOW + 14 * DAY);
    expect(fields.LinkStatus).toBe(LINK_STATUS.WAITING);
    expect(fields.ExpiresOn).toBe(new Date(NOW + 14 * DAY).toISOString());
    expect(JSON.parse(fields.Preset)).toEqual(signedValues);
  });

  it('lets the employee change again exactly what they could change before', () => {
    // IT fixed only the name; everything else was the employee's.
    const fields = reopenFields(signed({ preset: { employeeName: 'Amir Hakim' } }), NOW + DAY);
    const editable = JSON.parse(fields.Editable);
    expect(editable).toContain('position');
    expect(editable).toContain('serialNumbers');
    expect(editable).not.toContain('employeeName');
  });

  it('leaves the signed record and its signature where they are', () => {
    const fields = reopenFields(signed(), NOW + DAY);
    expect(fields).not.toHaveProperty('ChecklistId');
    expect(fields).not.toHaveProperty('SignatureFile');
    expect(fields).not.toHaveProperty('Submitted');
  });

  it('keeps IT\'s preset when the link was never signed', () => {
    expect(reopenFields(link({ status: LINK_STATUS.CANCELLED }), NOW + DAY)).not.toHaveProperty('Preset');
  });

  it('knows a reopened link from a fresh one', () => {
    expect(isReopened(link())).toBe(false);
    expect(isReopened(signed({ status: LINK_STATUS.WAITING }))).toBe(true);
  });
});

describe('IT editing a signed checklist', () => {
  const at = NOW;
  const by = 'IT Desk';

  it('writes the corrected values to the checklist row and marks it edited', () => {
    const plan = planEdit(signed(), { ...signedValues, serialNumbers: 'SN-2' }, { by, at });
    expect(plan.errors).toEqual({});
    expect(plan.checklist.SerialNumbers).toBe('SN-2');
    expect(plan.checklist.EditedAfterSigning).toBe(editNote(by, at));
    expect(plan.checklist.EditedAfterSigning).toMatch(/^Edited by IT Desk on \d\d\/10\/2026/);
  });

  it('never touches the signature, the form type or when it was signed', () => {
    const plan = planEdit(signed(), { ...signedValues, formMode: INDIVIDUAL, signature: 'data:image/png;base64,AA' }, { by, at });
    expect(plan.checklist).not.toHaveProperty('SignatureUrl');
    expect(plan.checklist).not.toHaveProperty('SubmissionDate');
    expect(plan.checklist).not.toHaveProperty('FormMode');
    expect(JSON.parse(plan.link.Submitted)).not.toHaveProperty('signature');
  });

  it('keeps the values as originally signed, the first time only', () => {
    const first = planEdit(signed(), { ...signedValues, position: 'Senior Engineer' }, { by, at });
    expect(JSON.parse(first.link.OriginalSubmitted)).toEqual(signedValues);

    const already = signed({ originalSubmitted: { ...signedValues, position: 'Intern' }, editedBy: by });
    const second = planEdit(already, signedValues, { by, at });
    expect(second.link).not.toHaveProperty('OriginalSubmitted');
  });

  it('records who edited and when on the link', () => {
    const plan = planEdit(signed(), signedValues, { by, at });
    expect(plan.link.EditedBy).toBe(by);
    expect(plan.link.EditedOn).toBe(new Date(at).toISOString());
  });

  it('still needs every required field', () => {
    const plan = planEdit(signed(), { ...signedValues, employeeNo: '' }, { by, at });
    expect(plan.errors.employeeNo).toBeTruthy();
    expect(plan.errors.signature).toBeUndefined();
  });

  it('refuses anything that is not signed', () => {
    expect(() => planEdit(link(), signedValues, { by, at })).toThrow(/signed/i);
  });
});

describe('the new link columns', () => {
  it('round-trip', () => {
    const item = {
      ...toLinkItem(signed()),
      EditedBy: 'IT Desk',
      EditedOn: '2026-10-05T04:00:00Z',
      OriginalSubmitted: JSON.stringify(signedValues),
    };
    const back = fromLinkItem(item);
    expect(back.editedBy).toBe('IT Desk');
    expect(back.editedOn).toBe('2026-10-05T04:00:00Z');
    expect(back.originalSubmitted).toEqual(signedValues);
  });
});

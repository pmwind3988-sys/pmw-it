import { describe, it, expect } from 'vitest';
import { IN, OUT, INDIVIDUAL, emptyChecklist } from '../checklistForm.js';
import { newLinkCode, isLinkCode, LINK_CODE_LENGTH } from './linkCode.js';
import {
  isBlankValue, employeeMayEdit, editableFields, linkState, cleanSubmission,
  mergeSubmission, cleanPreset, STALE_SIGNING_MS, MAX_SIGNATURE_CHARS,
} from './linkRules.js';
import { toLinkItem, fromLinkItem, LINK_STATUS } from './linkSchema.js';

const NOW = Date.parse('2026-10-05T04:00:00Z');
const DAY = 86400000;

const link = (overrides = {}) => ({
  id: 7,
  code: 'Ab3dE5gH9k',
  formMode: IN,
  preset: {
    employeeName: 'Amir Hakim',
    employeeNo: 'E-1042',
    position: '',
    entity: 'PMW',
    formDate: '2026-10-05',
    checkedItems: ['Laptop', 'Mouse'],
    items: [],
    serialNumbers: 'SN-1',
    otherRemarks: '',
  },
  editable: [],
  expiresOn: new Date(NOW + 3 * DAY).toISOString(),
  status: LINK_STATUS.WAITING,
  modified: new Date(NOW - DAY).toISOString(),
  ...overrides,
});

const signature = 'data:image/png;base64,iVBORw0KGgo=';

describe('the link code', () => {
  it('is ten letters and digits', () => {
    const code = newLinkCode();
    expect(code).toHaveLength(LINK_CODE_LENGTH);
    expect(code).toMatch(/^[A-Za-z0-9]{10}$/);
    expect(isLinkCode(code)).toBe(true);
  });

  it('differs every time', () => {
    const codes = new Set(Array.from({ length: 200 }, () => newLinkCode()));
    expect(codes.size).toBe(200);
  });

  it('skips the bytes that would favour the first letters of the alphabet', () => {
    // 248 and up are rejected (248 = 4 * 62), so 255 never becomes a letter.
    let call = 0;
    const random = (bytes) => {
      bytes.fill(call === 0 ? 255 : 0);
      call += 1;
      return bytes;
    };
    expect(newLinkCode(random)).toBe('AAAAAAAAAA');
  });

  it('refuses anything that is not a code before it reaches SharePoint', () => {
    expect(isLinkCode("x' or 1 eq 1")).toBe(false);
    expect(isLinkCode('short')).toBe(false);
    expect(isLinkCode(undefined)).toBe(false);
  });
});

describe('what the employee may change', () => {
  it('may always fill a field IT left blank', () => {
    expect(employeeMayEdit(link(), 'position')).toBe(true);
    expect(employeeMayEdit(link(), 'otherRemarks')).toBe(true);
  });

  it('may not change a filled field IT did not open', () => {
    expect(employeeMayEdit(link(), 'employeeName')).toBe(false);
    expect(employeeMayEdit(link(), 'checkedItems')).toBe(false);
  });

  it('may change a filled field IT opened', () => {
    expect(employeeMayEdit(link({ editable: ['serialNumbers'] }), 'serialNumbers')).toBe(true);
  });

  it('reads an item list with no named item as blank', () => {
    expect(isBlankValue('items', [{ item: '', quantity: 1 }])).toBe(true);
    expect(isBlankValue('items', [{ item: 'Mouse', quantity: 1 }])).toBe(false);
    expect(isBlankValue('checkedItems', [])).toBe(true);
  });

  it('lists only the fields this form type shows', () => {
    const fields = editableFields(link());
    expect(fields).toEqual(['position', 'otherRemarks']);
    expect(fields).not.toContain('items');
  });

  it('opens the item lines on an individual request IT left empty', () => {
    const request = link({ formMode: INDIVIDUAL, preset: { ...link().preset, items: [] } });
    expect(editableFields(request)).toContain('items');
  });
});

describe('the state of a link', () => {
  it('is open while waiting and in date', () => {
    expect(linkState(link(), NOW)).toBe('open');
  });

  it('is expired once past its date, however it was left', () => {
    expect(linkState(link({ expiresOn: new Date(NOW - 1).toISOString() }), NOW)).toBe('expired');
  });

  it('stays signed after the expiry date, so the copy can still be read', () => {
    const signed = link({ status: LINK_STATUS.SIGNED, expiresOn: new Date(NOW - DAY).toISOString() });
    expect(linkState(signed, NOW)).toBe('signed');
  });

  it('is cancelled when IT cancelled it', () => {
    expect(linkState(link({ status: LINK_STATUS.CANCELLED }), NOW)).toBe('cancelled');
  });

  it('is busy while a submission holds it, and open again if that submission died', () => {
    const fresh = link({ status: LINK_STATUS.SIGNING, modified: new Date(NOW - 1000).toISOString() });
    const stale = link({ status: LINK_STATUS.SIGNING, modified: new Date(NOW - STALE_SIGNING_MS - 1).toISOString() });
    expect(linkState(fresh, NOW)).toBe('busy');
    expect(linkState(stale, NOW)).toBe('open');
  });
});

describe('cleaning what the browser sent', () => {
  it('drops options the form never offered', () => {
    const clean = cleanSubmission({
      entity: 'NOPE',
      checkedItems: ['Laptop', 'Rocket'],
      items: [{ item: 'Mouse', quantity: '2' }, { item: 'Tank', quantity: 1 }],
    });
    expect(clean.entity).toBe('');
    expect(clean.checkedItems).toEqual(['Laptop']);
    expect(clean.items).toEqual([{ item: 'Mouse', quantity: 2 }]);
  });

  it('bounds quantities and lengths', () => {
    const clean = cleanSubmission({
      employeeName: 'x'.repeat(1000),
      otherRemarks: 'y'.repeat(10000),
      items: [{ item: 'Mouse', quantity: 5000 }, { item: 'Laptop', quantity: -3 }],
    });
    expect(clean.employeeName).toHaveLength(255);
    expect(clean.otherRemarks).toHaveLength(4000);
    expect(clean.items.map((row) => row.quantity)).toEqual([999, 1]);
  });

  it('accepts only a PNG signature of sensible size', () => {
    expect(cleanSubmission({ signature }).signature).toBe(signature);
    expect(cleanSubmission({ signature: 'data:text/html;base64,PGI+' }).signature).toBeNull();
    expect(cleanSubmission({ signature: `data:image/png;base64,${'A'.repeat(MAX_SIGNATURE_CHARS)}` }).signature).toBeNull();
  });

  it('keeps only a real calendar day', () => {
    expect(cleanSubmission({ formDate: '2026-10-05' }).formDate).toBe('2026-10-05');
    expect(cleanSubmission({ formDate: 'tomorrow' }).formDate).toBe('');
  });

  it('cleans IT\'s preset the same way, minus the signature', () => {
    const preset = cleanPreset({ ...emptyChecklist(), formMode: OUT, signature, employeeName: ' Siti ' });
    expect(preset.signature).toBeUndefined();
    expect(preset.formMode).toBeUndefined();
    expect(preset.employeeName).toBe('Siti');
  });
});

describe('merging the employee\'s answers onto IT\'s preset', () => {
  it('keeps what IT locked, whatever the browser sent', () => {
    const merged = mergeSubmission(link(), {
      employeeName: 'Someone Else',
      checkedItems: ['Laptop', 'Mouse', 'Monitor'],
      position: 'Engineer',
      signature,
    });
    expect(merged.employeeName).toBe('Amir Hakim');
    expect(merged.checkedItems).toEqual(['Laptop', 'Mouse']);
    expect(merged.position).toBe('Engineer');
    expect(merged.signature).toBe(signature);
  });

  it('takes the form type from the link, never from the browser', () => {
    expect(mergeSubmission(link(), { formMode: OUT }).formMode).toBe(IN);
  });

  it('takes an opened field from the employee', () => {
    const merged = mergeSubmission(link({ editable: ['serialNumbers'] }), { serialNumbers: 'SN-2' });
    expect(merged.serialNumbers).toBe('SN-2');
  });
});

describe('the link row', () => {
  it('round-trips through SharePoint fields', () => {
    const original = link();
    const item = toLinkItem(original, { createdByName: 'IT', createdByEmail: 'it@pmw.test' });
    expect(item.Title).toBe(original.code);
    expect(item.EmployeeName).toBe('Amir Hakim');

    const back = fromLinkItem({ ...item, id: '7', Modified: original.modified });
    expect(back.code).toBe(original.code);
    expect(back.preset).toEqual(original.preset);
    expect(back.editable).toEqual([]);
    expect(back.status).toBe(LINK_STATUS.WAITING);
    expect(back.id).toBe(7);
  });

  it('survives a row whose JSON somebody broke by hand', () => {
    const back = fromLinkItem({ Title: 'Ab3dE5gH9k', Preset: '{oops', Editable: 'nope' });
    expect(back.preset).toEqual({});
    expect(back.editable).toEqual([]);
  });
});

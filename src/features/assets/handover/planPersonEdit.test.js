import { describe, it, expect } from 'vitest';
import {
  planPersonEdit, personEditRefusal, personAt, normaliseEmail,
} from './planPersonEdit.js';
import { HANDOVER_STATUS } from './availability.js';

const handover = (over = {}) => ({
  id: 1,
  assetId: 10,
  assetKey: 'laptop-1',
  itemTitle: 'ThinkPad',
  personName: 'amir',
  personEmail: 'amir@pmw.com',
  personLogin: 'i:0#.f|membership|amir@pmw.com',
  quantity: 1,
  returnedQuantity: 0,
  handoverStatus: HANDOVER_STATUS.OUT,
  issueSignature: '/sites/it/Photos/signature-amir-1.png',
  ...over,
});

const asset = (over = {}) => ({
  id: 10,
  assetKey: 'laptop-1',
  title: 'ThinkPad',
  trackingMode: 'Tracked',
  quantity: 1,
  quantityOut: 1,
  assignedTo: 'amir',
  assignedToEmail: 'amir@pmw.com',
  ...over,
});

const edit = {
  from: 'amir@pmw.com',
  name: 'Amir Zahari',
  email: 'Amir@PMWgroup.com',
  login: '',
};

describe('planPersonEdit', () => {
  it('renames every row that person has, open and closed', () => {
    const rows = [
      handover(),
      handover({ id: 2, returnedQuantity: 1, handoverStatus: HANDOVER_STATUS.RETURNED }),
    ];

    const plan = planPersonEdit(rows, [asset()], edit);

    expect(plan.rows).toBe(2);
    // History follows the person. Correcting an email must not cut somebody's
    // record in half at the moment of the correction.
    expect(plan.handoverUpdates.map((update) => update.id)).toEqual([1, 2]);
    for (const update of plan.handoverUpdates) {
      expect(update.body.personName).toBe('Amir Zahari');
      expect(update.body.personEmail).toBe('amir@pmwgroup.com');
    }
  });

  it('leaves everybody else alone', () => {
    const rows = [handover(), handover({ id: 3, personEmail: 'siti@pmwgroup.com' })];

    const plan = planPersonEdit(rows, [], edit);

    expect(plan.handoverUpdates.map((update) => update.id)).toEqual([1]);
  });

  it('touches nothing about what is held', () => {
    const plan = planPersonEdit([handover()], [asset()], edit);
    const written = Object.keys(plan.handoverUpdates[0].body);

    // The whole promise of this screen: the person changes and the items do
    // not. A quantity, a date, a condition or a signature in this list would
    // be a rewrite of what happened.
    expect(written.sort()).toEqual(['personEmail', 'personLogin', 'personName']);
  });

  it('corrects the register copy on what they still hold', () => {
    const plan = planPersonEdit([handover()], [asset()], edit);

    expect(plan.assetUpdates).toEqual([{
      id: 10,
      assetKey: 'laptop-1',
      body: { assignedTo: 'Amir Zahari', assignedToEmail: 'amir@pmwgroup.com' },
    }]);
  });

  it('leaves the register alone for what they have already given back', () => {
    const rows = [handover({ returnedQuantity: 1, handoverStatus: HANDOVER_STATUS.RETURNED })];

    // The row is back on the shelf and names nobody. Writing a holder onto it
    // would put a laptop back in somebody's hands on paper.
    expect(planPersonEdit(rows, [asset({ assignedTo: '', assignedToEmail: '' })], edit)
      .assetUpdates).toEqual([]);
  });

  it('leaves a bulk row alone, which never named one holder', () => {
    const bulk = asset({ trackingMode: 'Bulk', assignedTo: '', assignedToEmail: '', quantity: 5 });

    expect(planPersonEdit([handover()], [bulk], edit).assetUpdates).toEqual([]);
  });

  it('drops a login that belonged to the old address', () => {
    const plan = planPersonEdit([handover()], [], edit);

    expect(plan.handoverUpdates[0].body.personLogin).toBe('');
  });

  it('keeps the login when only the spelling of a name was fixed', () => {
    const plan = planPersonEdit([handover()], [], {
      from: 'amir@pmw.com', name: 'Amir Zahari', email: 'amir@pmw.com',
    });

    expect('personLogin' in plan.handoverUpdates[0].body).toBe(false);
  });

  it('counts what is still out, for the screen to report', () => {
    const rows = [
      handover(),
      handover({ id: 2, returnedQuantity: 1, handoverStatus: HANDOVER_STATUS.RETURNED }),
    ];

    expect(planPersonEdit(rows, [], edit)).toMatchObject({ rows: 2, openLines: 1 });
  });
});

describe('personEditRefusal', () => {
  const current = { name: 'amir', email: 'amir@pmw.com' };

  it('accepts a real correction', () => {
    expect(personEditRefusal({ name: 'Amir Zahari', email: 'amir@pmwgroup.com' }, current))
      .toBeNull();
  });

  it('refuses an empty email, which is what everything is filed under', () => {
    expect(personEditRefusal({ name: 'Amir', email: '  ' }, current)).toMatch(/cannot be empty/);
  });

  it('refuses something that is not an address', () => {
    expect(personEditRefusal({ name: 'Amir', email: 'amir' }, current)).toMatch(/work email/);
    expect(personEditRefusal({ name: 'Amir', email: 'amir@' }, current)).toMatch(/work email/);
    expect(personEditRefusal({ name: 'Amir', email: 'a b@pmw.com' }, current)).toMatch(/space/);
  });

  it('refuses an empty name', () => {
    expect(personEditRefusal({ name: ' ', email: 'amir@pmwgroup.com' }, current))
      .toMatch(/cannot be empty/);
  });

  it('says so when nothing has actually changed', () => {
    expect(personEditRefusal({ name: 'amir', email: 'AMIR@pmw.com' }, current))
      .toMatch(/Nothing has been changed/);
  });
});

describe('personAt', () => {
  it('finds who else already answers to an address', () => {
    const rows = [
      handover({ id: 4, personEmail: 'amir@pmwgroup.com', personName: 'Amir Z' }),
      handover({
        id: 5,
        personEmail: 'amir@pmwgroup.com',
        returnedQuantity: 1,
        handoverStatus: HANDOVER_STATUS.RETURNED,
      }),
    ];

    expect(personAt(rows, 'Amir@PMWgroup.com')).toEqual({
      email: 'amir@pmwgroup.com', name: 'Amir Z', rows: 2, openLines: 1,
    });
  });

  it('is nothing when the address is new', () => {
    expect(personAt([handover()], 'siti@pmwgroup.com')).toBeNull();
    expect(personAt([handover()], '')).toBeNull();
  });
});

describe('normaliseEmail', () => {
  it('is how two spellings of one address become one identity', () => {
    expect(normaliseEmail('  Amir@PMW.com ')).toBe('amir@pmw.com');
    expect(normaliseEmail(null)).toBe('');
  });
});

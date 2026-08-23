import { describe, it, expect } from 'vitest';
import { toListItem, fromListItem, toUpdateItem } from './handoverSchema.js';
import { HANDOVER_KIND, HANDOVER_STATUS } from '../handover/availability.js';

const handover = {
  title: 'Amir — Dell Latitude 5540',
  handoverId: 'basket-1',
  assetKey: 'serial:DELL|CN0ABC',
  assetId: 7,
  itemTitle: 'Dell Latitude 5540',
  category: 'Laptop',
  personName: 'Amir',
  personEmail: 'amir@pmw.com',
  personLogin: 'i:0#.f|membership|amir@pmw.com',
  quantity: 1,
  returnedQuantity: 0,
  kind: HANDOVER_KIND.ISSUED,
  handoverStatus: HANDOVER_STATUS.OUT,
  issuedOn: 1755950400000,
  dueOn: null,
  returnedOn: null,
};

describe('toListItem', () => {
  it('sends the person and the item', () => {
    const item = toListItem(handover);

    expect(item.Title).toBe('Amir — Dell Latitude 5540');
    expect(item.PersonEmail).toBe('amir@pmw.com');
    expect(item.AssetId).toBe(7);
  });

  it('writes the readable time beside the instant', () => {
    const item = toListItem(handover);

    expect(item.IssuedOn).toBe(new Date(1755950400000).toISOString());
    expect(item.IssuedOnMYT).toMatch(/(AM|PM)/);
  });

  /**
   * An issued item legitimately has no due date and no return date. Unguarded,
   * `formatMYT` throws on those and every ordinary handover fails to write.
   */
  it('leaves the readable copy blank when there is no date', () => {
    const item = toListItem(handover);

    expect(item.DueOnMYT).toBe('');
    expect(item.ReturnedOnMYT).toBe('');
    expect('DueOn' in item).toBe(false);
  });

  /** Zero returned is a real answer and must survive. */
  it('sends a returned quantity of zero', () => {
    expect(toListItem(handover).ReturnedQuantity).toBe(0);
  });
});

describe('fromListItem', () => {
  it('round-trips', () => {
    const back = fromListItem({ Id: 3, ...toListItem(handover) });

    expect(back).toMatchObject({
      id: 3,
      personEmail: 'amir@pmw.com',
      quantity: 1,
      returnedQuantity: 0,
      kind: HANDOVER_KIND.ISSUED,
      handoverStatus: HANDOVER_STATUS.OUT,
    });
    expect(back.issuedOn).toBe(1755950400000);
    expect(back.dueOn).toBeNull();
  });

  /**
   * Both are summed and compared everywhere. A null would poison the
   * arithmetic in a way that reads as "nothing is out" rather than as an error.
   */
  it('defaults the quantities rather than leaving them null', () => {
    const back = fromListItem({ Id: 1 });

    expect(back.quantity).toBe(1);
    expect(back.returnedQuantity).toBe(0);
  });

  it('reads an absent date as null, not as NaN', () => {
    expect(fromListItem({ Id: 1 }).dueOn).toBeNull();
  });
});

describe('toUpdateItem', () => {
  it('sends only the fields a return touches', () => {
    const item = toUpdateItem({
      returnedQuantity: 2,
      handoverStatus: HANDOVER_STATUS.PARTLY,
      returnedOn: 1755950400000,
      returnCondition: 'Good',
    });

    expect(item.ReturnedQuantity).toBe(2);
    expect(item.HandoverStatus).toBe(HANDOVER_STATUS.PARTLY);
    expect(item.ReturnCondition).toBe('Good');
    expect('PersonEmail' in item).toBe(false);
    expect('Quantity' in item).toBe(false);
  });

  /** The readable copy has to move with the instant or the two disagree. */
  it('updates the readable return time alongside the instant', () => {
    expect(toUpdateItem({ returnedOn: 1755950400000 }).ReturnedOnMYT).toMatch(/(AM|PM)/);
  });

  it('sends a returned quantity of zero', () => {
    expect(toUpdateItem({ returnedQuantity: 0 }).ReturnedQuantity).toBe(0);
  });
});

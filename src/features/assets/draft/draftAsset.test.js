import { describe, it, expect } from 'vitest';
import {
  newDraft, draftFromCodes, setDraftField, swapSerialAndPart, draftIssues, isBlocked,
} from './draftAsset.js';
import { TRACKED, BULK } from '../assetKinds.js';

describe('newDraft', () => {
  it('takes its tracking mode from the category', () => {
    expect(newDraft({ category: 'Laptop' }).trackingMode).toBe(TRACKED);
    expect(newDraft({ category: 'Cable' }).trackingMode).toBe(BULK);
  });

  it('gives every draft an id of its own', () => {
    expect(newDraft().localId).not.toBe(newDraft().localId);
  });

  /** `undefined` means "inherit from the delivery"; '' would mean "no supplier". */
  it('leaves the purchase fields unset rather than blank', () => {
    const draft = newDraft();
    expect(draft.supplier).toBeUndefined();
    expect(draft.poNumber).toBeUndefined();
  });
});

describe('draftFromCodes', () => {
  it('fills the fields the codes identify and records what was guessed', () => {
    const draft = draftFromCodes([
      { rawValue: 'CN0ABC1234567', format: 'code_128' },
      { rawValue: 'P/N: 5UF44AA', format: 'code_128' },
    ]);

    expect(draft.serialNumber).toBe('CN0ABC1234567');
    expect(draft.partNumber).toBe('5UF44AA');
    expect(draft.guessed).toContain('serialNumber');
    expect(draft.guessed).not.toContain('partNumber');
  });

  /**
   * Nothing in a barcode says what the thing is, and the category decides
   * whether the row is tracked or counted — so guessing it would silently
   * decide the shape of the record.
   */
  it('does not guess a category', () => {
    expect(draftFromCodes([{ rawValue: 'CN0ABC123' }]).category).toBe('Other');
  });
});

describe('setDraftField', () => {
  it('stops a corrected field being a guess and marks it as hand-set', () => {
    const draft = draftFromCodes([{ rawValue: 'CN0ABC1234567' }]);
    const fixed = setDraftField(draft, 'serialNumber', 'CN0XYZ999');

    expect(fixed.serialNumber).toBe('CN0XYZ999');
    expect(fixed.guessed).not.toContain('serialNumber');
    expect(fixed.manualFields).toContain('serialNumber');
  });

  it('re-derives the tracking mode when the category changes', () => {
    const draft = newDraft({ category: 'Cable' });
    expect(setDraftField(draft, 'category', 'Laptop').trackingMode).toBe(TRACKED);
  });

  /** The override the spec promises: a serialised keyboard stays tracked. */
  it('leaves a hand-set tracking mode alone when the category changes', () => {
    const draft = setDraftField(newDraft({ category: 'Keyboard' }), 'trackingMode', TRACKED);
    const recategorised = setDraftField(draft, 'category', 'Mouse');

    expect(recategorised.trackingMode).toBe(TRACKED);
  });

  /** Twenty tracked units cannot share one serial number. */
  it('forces a tracked row back to a quantity of one', () => {
    const bulk = setDraftField(newDraft({ category: 'Cable' }), 'quantity', 20);
    expect(bulk.quantity).toBe(20);

    expect(setDraftField(bulk, 'trackingMode', TRACKED).quantity).toBe(1);
  });

  it('refuses a quantity that is not a positive whole number', () => {
    const draft = newDraft({ category: 'Cable' });

    expect(setDraftField(draft, 'quantity', '0').quantity).toBe(1);
    expect(setDraftField(draft, 'quantity', '-4').quantity).toBe(1);
    expect(setDraftField(draft, 'quantity', 'lots').quantity).toBe(1);
    expect(setDraftField(draft, 'quantity', '7.8').quantity).toBe(7);
  });

  /**
   * Ten monitors arriving together are one line reading ten, not ten rows.
   * The tracked-means-one rule is not broken by this — the row stops being
   * tracked, and each monitor's own serial goes to its unit record.
   */
  it('counts a tracked row by quantity once there is more than one of it', () => {
    const counted = setDraftField(newDraft({ category: 'Monitor' }), 'quantity', 10);

    expect(counted.quantity).toBe(10);
    expect(counted.trackingMode).toBe(BULK);
    expect(counted.manualFields).toContain('trackingMode');
  });

  it('leaves a single tracked item tracked', () => {
    const one = setDraftField(newDraft({ category: 'Monitor' }), 'quantity', 1);

    expect(one.quantity).toBe(1);
    expect(one.trackingMode).toBe(TRACKED);
  });

  /**
   * Lowering a count only ever HIDES units. Flipping the row back to tracked
   * here would pin it to one and take the other nine monitors' serials with it.
   */
  it('does not un-count a line when the quantity comes back down to one', () => {
    const counted = setDraftField(newDraft({ category: 'Monitor' }), 'quantity', 10);

    expect(setDraftField(counted, 'quantity', 1).trackingMode).toBe(BULK);
  });

  /** The flip was the person's decision, so re-categorising must not undo it. */
  it('keeps a counted line counted when the category changes', () => {
    const counted = setDraftField(newDraft({ category: 'Monitor' }), 'quantity', 10);

    expect(setDraftField(counted, 'category', 'Laptop').trackingMode).toBe(BULK);
  });

  it('does not list the same field as hand-set twice', () => {
    const once = setDraftField(newDraft(), 'model', 'A');
    const twice = setDraftField(once, 'model', 'B');

    expect(twice.manualFields.filter((name) => name === 'model')).toHaveLength(1);
  });
});

describe('draftIssues', () => {
  const good = () => newDraft({ category: 'Laptop', model: 'Latitude 5540', serialNumber: 'CN0ABC1' });

  it('is happy with a complete row', () => {
    expect(draftIssues(good())).toEqual([]);
  });

  it('asks for something to identify the thing by', () => {
    const issues = draftIssues(newDraft({ category: 'Cable' }));
    expect(issues.some((issue) => issue.field === 'model')).toBe(true);
  });

  it('warns — but does not block — when nothing will identify it next time', () => {
    const issues = draftIssues(newDraft({ category: 'Laptop', model: 'Latitude 5540' }));
    const warning = issues.find((issue) => issue.field === 'serialNumber');

    expect(warning).toBeDefined();
    expect(isBlocked(issues)).toBe(false);
  });

  it('blocks a sticker label already on something in the register', () => {
    const registerTags = new Map([['PMW-0142', { title: 'Dell P2422H', id: 7 }]]);
    const issues = draftIssues({ ...good(), assetTag: 'pmw-0142' }, { registerTags });

    expect(isBlocked(issues)).toBe(true);
    expect(issues[0].message).toContain('Dell P2422H');
    expect(issues[0].conflictWith).toBe(7);
  });

  it('blocks a sticker label used twice inside one batch', () => {
    const draft = { ...good(), assetTag: 'PMW-0142' };
    const batchTags = new Map([['PMW-0142', 'some-other-row']]);

    expect(isBlocked(draftIssues(draft, { batchTags }))).toBe(true);
  });

  it('does not count a row as clashing with itself', () => {
    const draft = { ...good(), assetTag: 'PMW-0142' };
    const batchTags = new Map([['PMW-0142', draft.localId]]);

    expect(isBlocked(draftIssues(draft, { batchTags }))).toBe(false);
  });
});

describe('swapSerialAndPart', () => {
  it('puts each code in the other field', () => {
    const draft = newDraft({ serialNumber: 'AAA111', partNumber: 'BBB222' });
    const after = swapSerialAndPart(draft);

    expect(after.serialNumber).toBe('BBB222');
    expect(after.partNumber).toBe('AAA111');
  });

  /** A correction outranks a future re-scan, the same as typing one would. */
  it('marks both as set by hand and neither as a guess', () => {
    const draft = newDraft({
      serialNumber: 'AAA111', partNumber: 'BBB222', guessed: ['serialNumber', 'partNumber'],
    });
    const after = swapSerialAndPart(draft);

    expect(after.guessed).toEqual([]);
    expect(after.manualFields).toEqual(expect.arrayContaining(['serialNumber', 'partNumber']));
  });

  it('does not list a field twice when it was already set by hand', () => {
    const draft = newDraft({
      serialNumber: 'AAA111', partNumber: 'BBB222', manualFields: ['serialNumber'],
    });

    expect(swapSerialAndPart(draft).manualFields.filter((f) => f === 'serialNumber')).toHaveLength(1);
  });
});

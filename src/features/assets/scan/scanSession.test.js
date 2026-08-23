import { describe, it, expect } from 'vitest';
import {
  createSession, seeCode, commitPool, discardPool, removeDraft, replaceDraft,
  SCAN_MODES, OUTCOMES,
} from './scanSession.js';

const code = (rawValue, format = 'code_128') => ({ rawValue, format });

/** Feed a run of codes through a session the way the camera loop does. */
function scan(session, codes) {
  return codes.reduce(
    (acc, entry) => {
      const step = seeCode(acc.session, entry);
      return { session: step.session, outcomes: [...acc.outcomes, step.outcome] };
    },
    { session, outcomes: [] },
  );
}

describe('many-items mode', () => {
  it('turns each new code into its own draft', () => {
    const { session } = scan(createSession(SCAN_MODES.MANY), [
      code('AAA111'), code('BBB222'), code('CCC333'),
    ]);

    expect(session.drafts).toHaveLength(3);
  });

  /**
   * The whole point of the seen list. A barcode held in frame is decoded
   * dozens of times a second, and a scanner that counts every one of those is
   * worse than no scanner at all.
   */
  it('never counts the same code twice, however many frames it appears in', () => {
    const { session, outcomes } = scan(createSession(SCAN_MODES.MANY), [
      code('AAA111'), code('AAA111'), code('AAA111'), code('AAA111'),
    ]);

    expect(session.drafts).toHaveLength(1);
    expect(outcomes).toEqual([
      OUTCOMES.ACCEPTED, OUTCOMES.DUPLICATE, OUTCOMES.DUPLICATE, OUTCOMES.DUPLICATE,
    ]);
  });

  /** Silence is indistinguishable from the scanner being broken. */
  it('answers a re-read rather than ignoring it', () => {
    const first = seeCode(createSession(SCAN_MODES.MANY), code('AAA111'));
    const again = seeCode(first.session, code('AAA111'));

    expect(again.outcome).toBe(OUTCOMES.DUPLICATE);
    expect(again.code).toBe('AAA111');
  });

  it('reports an empty read without disturbing the session', () => {
    const session = createSession(SCAN_MODES.MANY);
    const step = seeCode(session, code('   '));

    expect(step.outcome).toBe(OUTCOMES.EMPTY);
    expect(step.session).toBe(session);
  });

  it('classifies the single code it was given', () => {
    const { session } = scan(createSession(SCAN_MODES.MANY), [code('S/N: CN0ABC123')]);
    expect(session.drafts[0].serialNumber).toBe('CN0ABC123');
  });
});

describe('one-item mode', () => {
  it('pools every code on the box and decides nothing until it is confirmed', () => {
    const { session, outcomes } = scan(createSession(SCAN_MODES.ONE), [
      code('CN0ABC1234567'), code('P/N: 5UF44AA'), code('A4:BB:6D:1E:9F:02'),
    ]);

    expect(outcomes).toEqual([OUTCOMES.POOLED, OUTCOMES.POOLED, OUTCOMES.POOLED]);
    expect(session.drafts).toHaveLength(0);
    expect(session.pool).toHaveLength(3);
  });

  it('makes one item out of the pool, with each code in its right field', () => {
    const { session } = scan(createSession(SCAN_MODES.ONE), [
      code('CN0ABC1234567'), code('P/N: 5UF44AA'), code('A4:BB:6D:1E:9F:02'),
    ]);
    const { session: after, draft } = commitPool(session);

    expect(after.drafts).toHaveLength(1);
    expect(after.pool).toEqual([]);
    expect(draft.serialNumber).toBe('CN0ABC1234567');
    expect(draft.partNumber).toBe('5UF44AA');
    expect(draft.macAddress).toBe('A4:BB:6D:1E:9F:02');
  });

  /**
   * Carrying the box just recorded back into frame must not produce a second
   * item — which is why committing empties the pool but keeps `seen`.
   */
  it('still refuses a code from a box already confirmed', () => {
    const first = scan(createSession(SCAN_MODES.ONE), [code('CN0ABC1234567')]);
    const { session } = commitPool(first.session);
    const again = seeCode(session, code('CN0ABC1234567'));

    expect(again.outcome).toBe(OUTCOMES.DUPLICATE);
    expect(again.session.drafts).toHaveLength(1);
  });

  it('commits nothing when the pool is empty', () => {
    const session = createSession(SCAN_MODES.ONE);
    expect(commitPool(session).draft).toBeNull();
    expect(commitPool(session).session.drafts).toHaveLength(0);
  });

  it('carries overrides onto the committed draft', () => {
    const { session } = scan(createSession(SCAN_MODES.ONE), [code('CN0ABC1234567')]);
    const { draft } = commitPool(session, { category: 'Laptop' });

    expect(draft.category).toBe('Laptop');
    expect(draft.trackingMode).toBe('Tracked');
  });

  it('discards the box in frame without forgetting it was seen', () => {
    const { session } = scan(createSession(SCAN_MODES.ONE), [code('AAA111')]);
    const cleared = discardPool(session);

    expect(cleared.pool).toEqual([]);
    expect(seeCode(cleared, code('AAA111')).outcome).toBe(OUTCOMES.DUPLICATE);
  });
});

describe('editing the strip', () => {
  it('removes a draft', () => {
    const { session } = scan(createSession(SCAN_MODES.MANY), [code('AAA111'), code('BBB222')]);
    const after = removeDraft(session, session.drafts[0].localId);

    expect(after.drafts).toHaveLength(1);
    expect(after.drafts[0].serialNumber).toBe('BBB222');
  });

  it('replaces a draft in place, leaving the rest alone', () => {
    const { session } = scan(createSession(SCAN_MODES.MANY), [code('AAA111'), code('BBB222')]);
    const edited = { ...session.drafts[0], category: 'Laptop' };
    const after = replaceDraft(session, edited);

    expect(after.drafts[0].category).toBe('Laptop');
    expect(after.drafts[1]).toBe(session.drafts[1]);
  });
});

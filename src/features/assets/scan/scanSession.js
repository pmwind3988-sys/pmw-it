import { draftFromCodes } from '../draft/draftAsset.js';

/**
 * The camera's memory of what it has already read.
 *
 * The whole problem this solves: a barcode that stays in frame is decoded
 * thirty times a second, and a scanner that counts each of those is worse than
 * no scanner. Every accepted code is remembered for the length of the session,
 * so the same code can never be counted twice — and re-reading one is
 * ANSWERED rather than ignored, because silence and "the scanner is broken"
 * look identical from behind a phone.
 *
 * Pure and serialisable. The session survives the app being closed, so it
 * cannot hold a Set or a MediaStream.
 */

export const SCAN_MODES = { ONE: 'one', MANY: 'many' };

export const OUTCOMES = {
  ACCEPTED: 'accepted',
  POOLED: 'pooled',
  DUPLICATE: 'duplicate',
  EMPTY: 'empty',
};

export function createSession(mode = SCAN_MODES.MANY) {
  return {
    mode,
    // Codes accepted at any point in this session, in order. An array rather
    // than a Set so the session can be written to IndexedDB unchanged.
    seen: [],
    // ONE mode only: the codes read off the box currently in frame, waiting
    // to be classified together into a single item.
    pool: [],
    drafts: [],
  };
}

/**
 * One decoded code arriving from the camera.
 *
 * In MANY mode a new code becomes its own draft immediately — that mode is for
 * sweeping across a shelf, where each box shows one code. Where a box turns out
 * to carry several, the review grid can merge the rows back together.
 *
 * In ONE mode the code joins the pool and nothing is decided until the box is
 * confirmed, because which code is the serial can only be answered by looking
 * at all of them together.
 */
export function seeCode(session, code) {
  const raw = String(code?.rawValue ?? '').trim();
  if (!raw) return { session, outcome: OUTCOMES.EMPTY, code: raw };

  if (session.seen.includes(raw)) {
    return { session, outcome: OUTCOMES.DUPLICATE, code: raw };
  }

  const entry = { rawValue: raw, format: String(code?.format ?? '') };
  const seen = [...session.seen, raw];

  if (session.mode === SCAN_MODES.ONE) {
    return {
      session: { ...session, seen, pool: [...session.pool, entry] },
      outcome: OUTCOMES.POOLED,
      code: raw,
    };
  }

  const draft = draftFromCodes([entry]);
  return {
    session: { ...session, seen, drafts: [...session.drafts, draft] },
    outcome: OUTCOMES.ACCEPTED,
    code: raw,
    draft,
  };
}

/**
 * ONE mode: the box is done. Everything pooled becomes a single draft and the
 * pool empties for the next box — but `seen` is kept, so carrying the previous
 * box back into frame still does not produce a second item.
 */
export function commitPool(session, overrides = {}) {
  if (!session.pool.length) return { session, draft: null };

  const draft = draftFromCodes(session.pool, overrides);
  return {
    session: { ...session, pool: [], drafts: [...session.drafts, draft] },
    draft,
  };
}

/** Throw away the box in frame without losing the session's memory of it. */
export function discardPool(session) {
  return { ...session, pool: [] };
}

export function removeDraft(session, localId) {
  return { ...session, drafts: session.drafts.filter((draft) => draft.localId !== localId) };
}

export function replaceDraft(session, draft) {
  return {
    ...session,
    drafts: session.drafts.map((entry) => (entry.localId === draft.localId ? draft : entry)),
  };
}

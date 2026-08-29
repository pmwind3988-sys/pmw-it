import { isOpen } from './availability.js';

/**
 * Correcting who somebody is, without disturbing what they hold.
 *
 * Names are typed at a desk in a hurry and directories are not always
 * reachable, so a person arrives in the register as "amir" against
 * `amir@pmw.com` when the company is `pmwgroup.com`. Until now the only way to
 * fix that was to return everything and hand it all out again — which rewrites
 * the dates, loses the signatures, and tells a story about a laptop coming
 * back that never happened.
 *
 * So this changes the person and nothing else. Every handover row that person
 * has, open and closed, takes the new name and email; their history follows
 * them rather than being cut in half at the moment of the correction. The
 * counts, the dates, the conditions and the signatures are all left exactly as
 * they were — nothing here can move an item.
 *
 * The register's own copy of the holder is corrected too, but only on the rows
 * that actually name this person: a tracked laptop says who has it, and a box
 * of cables out with five people has never had a single holder to correct
 * (§4.2).
 */

export function normaliseEmail(value) {
  return String(value ?? '').trim().toLowerCase();
}

/**
 * Why this edit cannot be saved, or null when it can.
 *
 * The email is the identity everything per-person hangs off (§4.5), so an
 * empty or malformed one is refused rather than written — a row keyed on
 * nothing is a row nobody can ever find again.
 */
export function personEditRefusal({ email, name }, current = {}) {
  const next = normaliseEmail(email);

  if (!next) return 'A work email is what everything about a person is filed under. It cannot be empty.';
  if (!next.includes('@') || next.startsWith('@') || next.endsWith('@')) {
    return 'That does not look like a work email.';
  }
  if (/\s/.test(next)) return 'An email cannot contain a space.';
  if (!String(name ?? '').trim()) return 'A name is what the lists show. It cannot be empty.';

  if (next === normaliseEmail(current.email)
    && String(name).trim() === String(current.name ?? '').trim()) {
    return 'Nothing has been changed yet.';
  }

  return null;
}

/**
 * Whoever ELSE already answers to the new email.
 *
 * Not an error: it is usually the point. A typo that created a second Amir is
 * fixed by giving the wrong one the right email, and the two records become
 * one. But it is a big enough thing to happen silently, so it is reported and
 * the screen says so before anybody presses save.
 */
export function personAt(handovers, email) {
  const wanted = normaliseEmail(email);
  if (!wanted) return null;

  let name = '';
  let units = 0;
  let rows = 0;

  for (const row of handovers) {
    if (normaliseEmail(row.personEmail) !== wanted) continue;
    rows += 1;
    if (!name && row.personName) name = row.personName;
    if (isOpen(row)) units += 1;
  }

  return rows ? { email: wanted, name: name || wanted, rows, openLines: units } : null;
}

/**
 * `handovers` is every handover read; `register` is the asset rows. `from` is
 * the email the person is filed under now, and `name` / `email` / `login` are
 * what they should be.
 *
 * Answers with `{ handoverUpdates, assetUpdates, rows, openLines }` — the
 * writes to make, and what the person page should say happened.
 */
export function planPersonEdit(handovers, register, {
  from, name: nextName, email: nextEmail, login: nextLogin,
} = {}) {
  const was = normaliseEmail(from);
  const now = normaliseEmail(nextEmail);
  const name = String(nextName ?? '').trim();

  const handoverUpdates = [];
  const openAssetIds = new Set();
  let openLines = 0;

  for (const row of handovers) {
    if (normaliseEmail(row.personEmail) !== was) continue;

    const body = { personName: name, personEmail: now };

    // The login belongs to the address. Left alone when only the spelling of a
    // name was fixed; cleared when the email moved and no new one was given,
    // because a login pointing at the person this row is no longer about is
    // worse than a row that admits it does not know.
    if (now !== was) body.personLogin = nextLogin ?? '';

    handoverUpdates.push({ id: row.id, assetKey: row.assetKey, body });

    if (isOpen(row)) {
      openLines += 1;
      if (row.assetId != null) openAssetIds.add(row.assetId);
    }
  }

  // Only where the register itself names this person. A bulk row out with five
  // people carries no holder to correct, and writing one would invent a claim
  // that the row was never making.
  const assetUpdates = [];
  for (const asset of register) {
    if (!openAssetIds.has(asset.id)) continue;
    if (normaliseEmail(asset.assignedToEmail) !== was) continue;

    assetUpdates.push({
      id: asset.id,
      assetKey: asset.assetKey,
      body: { assignedTo: name, assignedToEmail: now },
    });
  }

  return {
    handoverUpdates,
    assetUpdates,
    rows: handoverUpdates.length,
    openLines,
  };
}

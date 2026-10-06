import { inFleet } from './status.js';
import { MATCH } from './matchIncoming.js';

/**
 * "Carmen already has CARMEN-HP. Is this a replacement?"
 *
 * Asked only when a person turns up on a machine that is NEW to them -- a
 * machine the register has never seen, or one it has seen under somebody
 * else. Re-scanning Carmen's own laptop while she also has a desktop must not
 * ask the question again on every import.
 */
export const ANSWERS = { STASH: 'stash', RETIRE: 'retire', KEEP: 'keep' };

const ownerKey = (owner) => String(owner ?? '').trim().replace(/\s+/g, ' ').toLowerCase();

export function replacementsFor(matches, existing) {
  const prompts = [];
  // A machine whose own report is in this batch is being re-scanned, not
  // replaced: stashing it as well would race its update.
  const rescanned = new Set(matches.map((match) => match.existing?.id).filter((id) => id !== null && id !== undefined));

  for (const match of matches) {
    if (match.kind === MATCH.DUPLICATE_SERIAL) continue;
    const owner = ownerKey(match.device.owner);
    if (!owner) continue;
    if (match.existing && ownerKey(match.existing.owner) === owner) continue;

    const selfId = match.existing?.id ?? null;
    for (const row of existing) {
      if (row.id === selfId || rescanned.has(row.id) || !inFleet(row) || ownerKey(row.owner) !== owner) continue;
      prompts.push({
        key: `${match.device.sourceFileName}|${row.id}`,
        sourceFileName: match.device.sourceFileName,
        incomingName: match.device.computerName,
        owner: match.device.owner,
        old: {
          id: row.id,
          computerName: row.computerName,
          deviceType: row.deviceType,
          status: row.status || 'In use',
          since: row.createdOn ?? row.importedOn ?? null,
        },
      });
    }
  }

  return prompts;
}

export function unanswered(prompts, answers, excluded = new Set()) {
  return prompts.filter((prompt) => !excluded.has(prompt.sourceFileName) && !answers[prompt.key]).length;
}

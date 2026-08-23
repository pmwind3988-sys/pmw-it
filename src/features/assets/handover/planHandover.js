import { TRACKED } from '../assetKinds.js';
import { resolveLines } from './basket.js';
import {
  available, owned, out, HANDOVER_KIND, HANDOVER_STATUS,
} from './availability.js';

/**
 * What handing a basket over is going to write, decided before a request is
 * sent.
 *
 * Pure, for the same reason `planSave` is: every refusal is a rule somebody
 * will eventually disagree with, and a rule that can only be exercised through
 * a network is a rule nobody checks.
 */

/**
 * Two lines for the same box of cables in one basket.
 *
 * They have to be added up BEFORE the availability check, or two lines of three
 * each pass individually against a stock of five and hand out six. Tracked
 * items cannot collide this way — `hasAsset` stops the same laptop being added
 * twice — but the check costs nothing and the alternative is a silent overdraw.
 */
export function coalesceLines(lines) {
  const byAsset = new Map();

  for (const line of lines) {
    const existing = byAsset.get(line.assetId);
    if (!existing) {
      byAsset.set(line.assetId, { ...line });
      continue;
    }

    if (line.trackingMode !== TRACKED) existing.quantity += line.quantity ?? 0;
    // The stricter of the two wins: a line asked for on loan is on loan.
    if (line.kind === HANDOVER_KIND.BORROWED) {
      existing.kind = HANDOVER_KIND.BORROWED;
      existing.dueOn = existing.dueOn ?? line.dueOn;
    }
    existing.remarks = [existing.remarks, line.remarks].filter(Boolean).join(' — ');
  }

  return [...byAsset.values()];
}

/**
 * `register` is the rows as SharePoint currently has them, re-read immediately
 * before the writes — which is what stops two people issuing the same laptop
 * from two phones (§8).
 *
 * Returns `{ handovers, assetUpdates, blocked }`. A blocked line never stops
 * the others: one refusal is one line's problem.
 */
export function planHandover(basket, register, { issuedOn = Date.now(), issuedBy = '' } = {}) {
  const byId = new Map(register.map((asset) => [asset.id, asset]));

  const handovers = [];
  const assetUpdates = [];
  const blocked = [];

  const person = basket.person ?? {};

  for (const line of coalesceLines(resolveLines(basket))) {
    const asset = byId.get(line.assetId);

    if (!asset) {
      blocked.push({ line, reason: 'That item is no longer in the register.' });
      continue;
    }

    // A serialised thing recorded in two places at once is the failure that
    // makes people stop believing the register, so it is refused outright
    // rather than warned about.
    if (asset.trackingMode === TRACKED && available(asset) < 1) {
      blocked.push({
        line,
        reason: asset.assignedTo
          ? `${asset.assignedTo} already has this. Take it back first.`
          : 'This is already out with somebody.',
        conflictWith: asset.id,
      });
      continue;
    }

    const wanted = asset.trackingMode === TRACKED ? 1 : (line.quantity ?? 0);
    const free = available(asset);

    if (wanted > free) {
      blocked.push({
        line,
        reason: `Only ${free} of ${owned(asset)} available.`,
      });
      continue;
    }

    handovers.push({
      handoverId: basket.id,
      assetKey: asset.assetKey,
      assetId: asset.id,
      itemTitle: asset.title ?? line.itemTitle,
      category: asset.category ?? '',
      personName: person.name ?? '',
      personEmail: person.email ?? '',
      personLogin: person.login ?? '',
      quantity: wanted,
      returnedQuantity: 0,
      kind: line.kind,
      handoverStatus: HANDOVER_STATUS.OUT,
      issuedOn,
      dueOn: line.kind === HANDOVER_KIND.BORROWED ? (line.dueOn ?? null) : null,
      returnedOn: null,
      returnCondition: '',
      issuedBy,
      returnedBy: '',
      remarks: line.remarks ?? '',
      title: `${person.name || person.email || 'Someone'} — ${asset.title ?? ''}`,
    });

    assetUpdates.push({
      id: asset.id,
      assetKey: asset.assetKey,
      body: {
        quantityOut: out(asset) + wanted,
        // Only a tracked row can name one holder. A box of cables can be with
        // five people at once and there is no honest single value, so those
        // fields stay empty and `QuantityOut` carries the answer (§4.2).
        ...(asset.trackingMode === TRACKED ? {
          status: line.kind === HANDOVER_KIND.BORROWED ? 'Borrowed' : 'Assigned',
          handoverKind: line.kind,
          assignedTo: person.name ?? '',
          assignedToEmail: person.email ?? '',
          assignedOn: issuedOn,
          dueOn: line.kind === HANDOVER_KIND.BORROWED ? (line.dueOn ?? null) : null,
        } : {}),
      },
    });
  }

  return { handovers, assetUpdates, blocked };
}

/**
 * What a line would be refused for right now, for showing against it as it is
 * added rather than at the end. Same rules as `planHandover`, asked one line at
 * a time — the plan stays the authority, and this is what the basket screen
 * calls to stay honest while it is being filled.
 */
export function lineRefusal(line, asset, basket) {
  if (!asset) return 'That item is no longer in the register.';

  const already = (basket?.lines ?? [])
    .filter((entry) => entry.assetId === asset.id && entry.lineId !== line.lineId)
    .reduce((sum, entry) => sum + (entry.quantity ?? 0), 0);

  if (asset.trackingMode === TRACKED) {
    if (available(asset) < 1) {
      return asset.assignedTo
        ? `${asset.assignedTo} already has this. Take it back first.`
        : 'This is already out with somebody.';
    }
    return null;
  }

  const wanted = (line.quantity ?? 0) + already;
  if (wanted > available(asset)) {
    return `Only ${available(asset)} of ${owned(asset)} available.`;
  }

  return null;
}

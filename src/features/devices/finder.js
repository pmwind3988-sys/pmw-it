import { hardwareIdsOf } from './derive/deriveHardwareIds.js';
import { softwareOf } from './derive/deriveSoftware.js';
import { programsMatching } from './fleetFacts.js';

/**
 * "Find a machine" — the search pill on /devices. Pure: it reads what a person
 * would type (name, owner, serials, product number, AnyDesk ID, memory and
 * drive serials or part numbers, installed programs), ranks the matches so the
 * machine they mean is first, and says WHY each one matched.
 *
 * Searches every saved machine, retired and spare ones included: a laptop
 * handed back last week is exactly what somebody goes looking for.
 */

const MIN_QUERY = 2;

/** AnyDesk IDs are read out with spaces ("123 456 789"); typed without them. */
const compact = (value) => String(value ?? '').replace(/\s+/g, '').toLowerCase();
const plain = (value) => String(value ?? '').trim().toLowerCase();

/**
 * Tiers, best first. A name match always beats a program match, whatever the
 * wording: the tier is the first thing compared, the closeness of the match
 * second.
 */
const TIER = { name: 0, owner: 1, serial: 2, anydesk: 3, part: 4, program: 5 };

/** 0 = exact, 1 = starts with the query, 2 = contains it, null = no match. */
function closeness(value, text) {
  if (!value) return null;
  if (value === text) return 0;
  if (value.startsWith(text)) return 1;
  if (value.includes(text)) return 2;
  return null;
}

const slotNote = (slot) => (slot ? ` · slot ${slot}` : '');

/** Every place a machine can be found, as { tier, label, value, suffix, squash }. */
function fieldsOf(device) {
  const hardware = hardwareIdsOf(device);
  const fields = [
    { tier: TIER.name, label: 'Computer name', value: device.computerName },
    // Somebody holding a laptop with the old name on a sticker still finds it.
    ...(device.previousNames ?? []).map((name) => ({ tier: TIER.name, label: 'Was named', value: name })),
    { tier: TIER.owner, label: 'Owner', value: device.owner },
    { tier: TIER.serial, label: 'Serial number', value: device.serialNumber },
    { tier: TIER.serial, label: 'Product number', value: device.productNumber },
    { tier: TIER.serial, label: 'Motherboard serial', value: device.motherboardSerial },
    { tier: TIER.anydesk, label: 'AnyDesk', value: device.anydeskId, squash: true },
  ];
  for (const module of hardware.memoryModules) {
    fields.push({ tier: TIER.part, label: 'Memory serial', value: module.serial, suffix: slotNote(module.slot) });
    fields.push({ tier: TIER.part, label: 'Memory part', value: module.partNumber, suffix: slotNote(module.slot) });
  }
  for (const drive of hardware.driveSerials) {
    fields.push({ tier: TIER.part, label: 'Drive serial', value: drive.serial });
  }
  for (const program of softwareOf(device)) {
    fields.push({ tier: TIER.program, label: 'Has', value: program.name });
  }
  return fields;
}

/**
 * The machines matching `query`, best first, with the reason for each. At most
 * `limit`. Ties keep the order the name sorts into, so the list does not
 * shuffle as you type.
 *
 * A reason is { label, value, suffix, at, length }: the text reads
 * `label value suffix`, and `at`/`length` mark the part of `value` that matched
 * so the view can bold it (-1 when the match was squashed, e.g. an AnyDesk ID).
 */
export function findMatches(devices, query, limit = 8) {
  const text = plain(query);
  const squashed = compact(query);
  if (squashed.length < MIN_QUERY || text.length < MIN_QUERY) return [];

  const ranked = [];
  for (const device of devices ?? []) {
    let best = null;
    for (const field of fieldsOf(device)) {
      const raw = String(field.value ?? '');
      if (!raw.trim()) continue;
      let closest;
      let at = -1;
      if (field.squash) {
        closest = closeness(compact(raw), squashed);
      } else {
        const lowered = plain(raw);
        closest = closeness(lowered, text);
        at = closest === null ? -1 : lowered.indexOf(text);
      }
      if (closest === null) continue;
      const key = field.tier * 3 + closest;
      if (!best || key < best.key) {
        best = {
          key,
          reason: {
            label: field.label,
            value: raw,
            suffix: field.suffix ?? '',
            at,
            length: at === -1 ? 0 : text.length,
          },
        };
      }
    }
    if (best) ranked.push({ device, ...best });
  }

  ranked.sort((a, b) => (a.key - b.key)
    || String(a.device.computerName ?? '').localeCompare(String(b.device.computerName ?? '')));

  return ranked.slice(0, limit).map(({ device, reason }) => ({ device, reason }));
}

/** The machines matching `query`, best first, at most `limit` of them. */
export function findMachines(devices, query, limit = 8) {
  return findMatches(devices, query, limit).map((match) => match.device);
}

/** How many machines have an installed program whose name contains `query`. */
export function machinesWithProgram(devices, query) {
  return (devices ?? []).filter((device) => programsMatching(device, query).length > 0).length;
}

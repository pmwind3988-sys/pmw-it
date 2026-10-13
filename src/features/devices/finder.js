/**
 * "Find a machine" — the search pill on /devices. Pure: it reads the fields a
 * person would actually type (name, owner, serial, AnyDesk ID) and ranks the
 * matches so the machine they mean is the first one offered.
 *
 * Searches every saved machine, retired and spare ones included: a laptop
 * handed back last week is exactly what somebody goes looking for.
 */

const MIN_QUERY = 2;

/** AnyDesk IDs are read out with spaces ("123 456 789"); typed without them. */
const compact = (value) => String(value ?? '').replace(/\s+/g, '').toLowerCase();
const plain = (value) => String(value ?? '').trim().toLowerCase();

/**
 * 0 = exact name or serial, 1 = starts with the query, 2 = contains it,
 * null = no match. Lower ranks first.
 */
function rankOf(device, query, squashed) {
  const name = plain(device.computerName);
  const serial = plain(device.serialNumber);
  const owner = plain(device.owner);
  const anydesk = compact(device.anydeskId);

  if (name === query || serial === query || (squashed && anydesk === squashed)) return 0;
  if (
    name.startsWith(query)
    || serial.startsWith(query)
    || owner.startsWith(query)
    || (squashed && anydesk.startsWith(squashed))
  ) return 1;
  if (
    name.includes(query)
    || serial.includes(query)
    || owner.includes(query)
    || (squashed && anydesk.includes(squashed))
  ) return 2;
  return null;
}

/**
 * The machines matching `query`, best first, at most `limit` of them. Ties
 * keep the order the name sorts into, so the list does not shuffle as you type.
 */
export function findMachines(devices, query, limit = 8) {
  const text = plain(query);
  const squashed = compact(query);
  if (squashed.length < MIN_QUERY || text.length < MIN_QUERY) return [];

  const ranked = [];
  for (const device of devices ?? []) {
    const rank = rankOf(device, text, squashed);
    if (rank !== null) ranked.push({ device, rank });
  }

  ranked.sort((a, b) => (a.rank - b.rank)
    || String(a.device.computerName ?? '').localeCompare(String(b.device.computerName ?? '')));

  return ranked.slice(0, limit).map((entry) => entry.device);
}

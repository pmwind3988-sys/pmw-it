/**
 * Pure: the names a machine used to answer to, oldest first.
 *
 * Kept on the machine's own row (`PreviousNames`) as well as in the change log,
 * so the List, the machine page and "Find a machine" can say "was AMIR-HP"
 * without reading the history, and a SharePoint view of the list shows it too.
 *
 * Case-insensitive, like every other name comparison in the register. Renaming
 * back to a name it once had takes that name OFF the list -- it is its name
 * again, not a former one.
 */
const key = (name) => String(name ?? '').trim().toLowerCase();

export function withPreviousName(previous, oldName, newName) {
  const old = String(oldName ?? '').trim();
  const list = (previous ?? []).map((name) => String(name).trim()).filter(Boolean);
  const kept = list.filter((name) => key(name) !== key(old) && key(name) !== key(newName));
  if (old && key(old) !== key(newName)) kept.push(old);
  return kept;
}

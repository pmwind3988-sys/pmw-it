/**
 * The Details tab's "Find a field": keep the groups whose fields have the
 * text in their label or their value, case-insensitively. An empty search
 * returns the groups untouched, and a group with no match is dropped so the
 * page never shows a card of nothing.
 */
export function matchingGroups(groups, query, valueOf) {
  const needle = String(query ?? '').trim().toLowerCase();
  if (!needle) return groups;
  return groups
    .map((group) => ({
      ...group,
      fields: group.fields.filter((field) => (
        String(field.label ?? '').toLowerCase().includes(needle)
        || String(valueOf(field) ?? '').toLowerCase().includes(needle)
      )),
    }))
    .filter((group) => group.fields.length > 0);
}

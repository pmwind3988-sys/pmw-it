// The order to take into a meeting -- spec §6.8.
//
// Distinct respondents leads, severity only scales it. One person
// writing five furious sentences must not outrank five people each
// writing one calm one: the first is an individual's bad week, the
// second is a process problem. Scoring on fragment count would get that
// exactly backwards, which is why `count` is carried for display and
// never enters the score.

export function rankIssues(groups, { pinned = [], suppressed = [] } = {}) {
  const pinOrder = new Map(pinned.map((id, i) => [id, i]));
  const suppressedSet = new Set(suppressed);

  const scored = (groups ?? []).map((group) => ({
    ...group,
    score: group.respondents * (1 + (group.meanSeverity ?? 0)),
    pinned: pinOrder.has(group.id),
    suppressed: suppressedSet.has(group.id),
  }));

  return scored.sort((a, b) => {
    // Suppressed sinks, pinned floats. Suppression is checked first so a
    // pinned item the user later hid does not float above live ones.
    if (a.suppressed !== b.suppressed) return a.suppressed ? 1 : -1;
    if (a.pinned !== b.pinned) return a.pinned ? -1 : 1;
    if (a.pinned && b.pinned) return pinOrder.get(a.id) - pinOrder.get(b.id);

    if (b.score !== a.score) return b.score - a.score;
    if ((b.meanSeverity ?? 0) !== (a.meanSeverity ?? 0)) {
      return (b.meanSeverity ?? 0) - (a.meanSeverity ?? 0);
    }
    // Alphabetical last, so two identical groups do not swap places
    // between renders for no reason.
    return String(a.label).localeCompare(String(b.label));
  });
}

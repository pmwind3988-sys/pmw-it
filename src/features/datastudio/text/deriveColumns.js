// The analysis, as ordinary spreadsheet columns -- spec §9.
//
// This is the payoff of the whole design. The results are appended to
// the raw grid as five more columns and everything downstream -- the
// profiler, the cleaner, the chart canvas, the tile editor, the filter
// bar, cross-filtering, saved dashboards, PNG export -- consumes them
// with no change whatsoever. The analysis adds data; it does not add a
// second charting system.
//
// `Issue categories` is deliberately semicolon-joined rather than one
// column per bucket. That makes it a multi column (see the profiler),
// so a chart counts it by OPTION, and one respondent who raised three
// kinds of problem is counted under all three instead of becoming their
// own private category.

export const NO_ISSUE_LABEL = 'No issue raised';
export const MULTI_SEPARATOR = ';';

export const DERIVED_HEADERS = [
  'Issue category',
  'Issue categories',
  'Theme',
  'Issue count',
  'Severity',
];

export function deriveColumns(analysis, rowCount) {
  const bucketLabel = new Map((analysis?.buckets ?? []).map((b) => [b.id, b.label]));
  const themeName = new Map((analysis?.themes ?? []).map((t) => [t.id, t.name]));

  const byRow = new Map();
  for (const fragment of analysis?.fragments ?? []) {
    if (fragment.noise) continue;
    if (!byRow.has(fragment.row)) byRow.set(fragment.row, []);
    byRow.get(fragment.row).push(fragment);
  }

  const primary = [];
  const all = [];
  const theme = [];
  const counts = [];
  const severities = [];

  for (let row = 0; row < rowCount; row++) {
    const fragments = byRow.get(row) ?? [];

    if (fragments.length === 0) {
      primary.push(NO_ISSUE_LABEL);
      all.push(NO_ISSUE_LABEL);
      theme.push(NO_ISSUE_LABEL);
      counts.push(0);
      severities.push(0);
      continue;
    }

    // The worst one speaks for the row. Picking the first would make the
    // headline category depend on the order somebody wrote their
    // sentences in.
    let worst = fragments[0];
    for (const fragment of fragments) {
      if ((fragment.severity ?? 0) > (worst.severity ?? 0)) worst = fragment;
    }

    const labels = [];
    for (const fragment of fragments) {
      const label = bucketLabel.get(fragment.bucketId);
      if (label && !labels.includes(label)) labels.push(label);
    }

    primary.push(bucketLabel.get(worst.bucketId) ?? NO_ISSUE_LABEL);
    all.push(labels.join(MULTI_SEPARATOR));
    theme.push(themeName.get(worst.themeId) ?? NO_ISSUE_LABEL);
    counts.push(fragments.length);
    // Whole numbers out of a hundred: a 0-1 float axis reads as a
    // proportion of something, which severity is not.
    severities.push(Math.round((worst.severity ?? 0) * 100));
  }

  return {
    headers: DERIVED_HEADERS,
    columns: [primary, all, theme, counts, severities],
  };
}

/**
 * What the profiler should be TOLD about these columns rather than left
 * to infer.
 *
 * `Issue categories` is multi-valued by construction. The multi-select
 * heuristic requires most cells to contain a separator, and on a real
 * survey most respondents raise a single category -- so it correctly
 * declines, and the column would chart by combination instead of by
 * option. There is nothing to infer here: this module wrote the column.
 */
export const DERIVED_OVERRIDES = {
  'Issue categories': { type: 'multi' },
};

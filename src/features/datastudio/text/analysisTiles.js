// A dashboard built from the text analysis itself.
//
// `deriveColumns.js` turns the analysis into five ordinary columns; this
// turns those five columns into the tiles that chart them. The split is
// the point: this file knows what the analysis MEANS (severity is worth
// averaging, categories are worth counting by option), and nothing
// downstream has to know the tiles came from an analysis rather than
// from a spreadsheet.
//
// The suggester in `suggest/suggestCharts.js` is not reused here, and
// deliberately. It ranks columns by SHAPE -- how many distinct values,
// how much variance -- because on an unknown sheet shape is the only
// evidence there is. Here the meaning of every column is known exactly,
// so guessing at it would be strictly worse: it would rank "Issue
// count" as a measure to sum against whatever dimension scored highest,
// which is not the question anybody asks of a pile of complaints.

import { DERIVED_HEADERS } from './deriveColumns.js';

// Stable ids, not generated ones. Pressing the button twice must
// REPLACE this dashboard rather than stack a second copy of it on top,
// and the ids are how the canvas tells the two apart.
export const ANALYSIS_TILE_PREFIX = 'txt_';

function tile(id, title, chart, encoding, extra = {}) {
  return {
    id: `${ANALYSIS_TILE_PREFIX}${id}`,
    title,
    chart,
    encoding: { x: null, y: [], series: null, ...encoding },
    sort: { by: 'y', dir: 'desc' },
    limit: 10,
    size: 'M',
    stacked: false,
    respondsToFilters: true,
    ...extra,
  };
}

export function isAnalysisTile(t) {
  return typeof t?.id === 'string' && t.id.startsWith(ANALYSIS_TILE_PREFIX);
}

/**
 * The tiles that chart an analysis, in reading order.
 *
 * `headers` is the grid's header list. A tile is only offered when the
 * column it needs is actually there, so a future analysis that stops
 * producing one of the five columns loses that tile rather than putting
 * a broken one on the canvas.
 *
 * Note what is NOT filtered out: "No issue raised". It is a real answer
 * -- most respondents to most surveys have no complaint -- and hiding
 * it would make a category raised by four people out of two hundred
 * look like the whole picture.
 */
export function analysisTiles(headers = DERIVED_HEADERS) {
  const has = (name) => headers.includes(name);
  const tiles = [];

  if (has('Issue count')) {
    tiles.push(tile(
      'kpi_issues', 'Issues raised', 'kpi',
      { y: [{ column: 'Issue count', agg: 'sum' }] },
      { size: 'S' },
    ));
  }

  tiles.push(tile(
    'kpi_people', 'People who answered', 'kpi',
    { y: [{ column: null, agg: 'count' }] },
    { size: 'S' },
  ));

  if (has('Severity')) {
    tiles.push(tile(
      'kpi_severity', 'Average severity out of 100', 'kpi',
      { y: [{ column: 'Severity', agg: 'avg' }] },
      { size: 'S' },
    ));
  }

  // The multi column, not the single one: a person who raised three
  // kinds of problem belongs in all three bars, and counting the
  // headline category alone would under-report every category but the
  // worst one on that row.
  if (has('Issue categories')) {
    tiles.push(tile(
      'categories', 'People by issue category', 'bar',
      { x: { column: 'Issue categories' }, y: [{ column: null, agg: 'count' }] },
      { size: 'L', limit: 12 },
    ));
  }

  if (has('Theme')) {
    tiles.push(tile(
      'themes', 'People by theme', 'bar',
      { x: { column: 'Theme' }, y: [{ column: null, agg: 'count' }] },
      { limit: 12 },
    ));
  }

  // Averaged, never summed. Summing severity would rank a category that
  // twenty people mentioned mildly above one that three people are
  // furious about, which is the opposite of what the number is for --
  // and how loud a category is already has its own chart above.
  if (has('Severity') && has('Issue category')) {
    tiles.push(tile(
      'severity', 'How severe each category is', 'bar',
      { x: { column: 'Issue category' }, y: [{ column: 'Severity', agg: 'avg' }] },
      { limit: 12 },
    ));
  }

  return tiles;
}

/**
 * The analysis dashboard, placed above whatever was already on the
 * canvas.
 *
 * Existing tiles are kept, not replaced. Somebody who built charts of
 * their own before running the analysis should not lose them to a
 * button labelled "build me a dashboard"; the analysis tiles simply
 * lead, because they are what was just asked for.
 */
export function withAnalysisTiles(existing = [], headers = DERIVED_HEADERS) {
  return [...analysisTiles(headers), ...existing.filter((t) => !isAnalysisTile(t))];
}

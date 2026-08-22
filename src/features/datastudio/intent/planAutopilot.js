// One decision per import: what to hide, what to chart, what to read.
//
// This is the only place that turns "a file landed" into "a dashboard is
// on screen". Everything it needs has already been measured -- the
// profile from `profile/`, the prose columns from `text/`, the title
// from `fileIntent` -- so it stays a pure function over plain data and
// the provider does nothing but carry out the result.
//
// The plan is also the disclosure. `hidden` and `analyseColumn` are what
// the brief card renders back to the user, which is why each hidden
// column carries its reason rather than just its name: a tool that
// silently drops columns is a tool nobody can trust with a sheet they
// have not read themselves.

import { readFileIntent, relevanceOf } from './fileIntent.js';
import { detectAdminColumns } from './adminColumns.js';

// Roles that still draw something once the hiding is done. If nothing
// survives, the lexicon has misread the sheet and hiding is abandoned
// wholesale -- an empty canvas is a worse answer than a cluttered one.
const CHARTABLE = new Set(['measure', 'dimension', 'temporal']);

/**
 * Which written-answer column to read first.
 *
 * `detectTextColumns` already sorts by how much people wrote, which is
 * the right default. The title only overrules it when a column name
 * genuinely echoes the title -- on a sheet asking about issues, the
 * column called "Biggest issue you face" is the question even if
 * respondents wrote more in the one below it.
 */
export function pickAnalyseColumn(textColumns = [], keywords = []) {
  if (textColumns.length === 0) return null;

  let best = textColumns[0];
  let bestScore = relevanceOf(best.name, keywords);

  for (const column of textColumns.slice(1)) {
    const score = relevanceOf(column.name, keywords);
    // Strictly greater: a tie leaves the longest-written column in
    // front, which is the order it arrived in.
    if (score > bestScore) {
      best = column;
      bestScore = score;
    }
  }

  return best.name;
}

/**
 * Read a freshly parsed sheet into a plan for the canvas.
 *
 * `autoAnalyse` is separate from `analyseColumn` on purpose. A column
 * worth offering is not the same as a column worth spending a 23MB model
 * download on unprompted: the analysis starts by itself when the title
 * says the writing IS the data (a pain-point or feedback survey), and
 * otherwise waits behind a button in the brief.
 */
export function planAutopilot({
  fileName = '', sheetName = '', profile = null, textColumns = [],
}) {
  const intent = readFileIntent(fileName, sheetName);
  const columns = profile?.columns ?? [];

  // A prose column is never bookkeeping, whatever its header says --
  // "Any other comments" is the answer, not the envelope.
  const prose = new Set(textColumns.map((c) => c.name));
  let hidden = detectAdminColumns(profile).filter((c) => !prose.has(c.name));

  const hiddenNames = new Set(hidden.map((c) => c.name));
  const survives = columns.some(
    (c) => CHARTABLE.has(c.role) && !hiddenNames.has(c.name),
  );
  if (!survives) hidden = [];

  const analyseColumn = pickAnalyseColumn(textColumns, intent.keywords);

  return {
    intent,
    hidden,
    // Handed to `suggestCharts` so the columns the title names get the
    // first chart slots.
    focus: intent.keywords,
    analyseColumn,
    autoAnalyse: Boolean(analyseColumn) && intent.textFirst,
  };
}

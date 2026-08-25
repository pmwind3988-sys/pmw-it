// Who answered, when, and from where — pulled to the front of a response.
//
// A Microsoft Forms export puts its bookkeeping down the left: an id, a
// start time, a completion time, an email address, a name. The charts
// deliberately ignore all of it (see `intent/adminColumns.js`), but the
// list of responses is the opposite case — somebody who has just tapped
// a bar wants to know WHO is behind it, and a table that opens on "Id,
// Start time, Completion time" has spent its first three columns saying
// nothing.
//
// So this splits a response in two: the three fields that identify a
// respondent, and the questions they answered. Matched on the HEADER,
// because shape cannot tell a submission timestamp from an incident
// date — both are datetimes with one value per row.

import { normalizeHeader } from '../intent/adminColumns.js';

/**
 * The three identity fields, in the order they are shown.
 *
 * Order matters and is not the sheet's: email first because it is the
 * one field a reader recognises a person by, then when they answered,
 * then which part of the business they are in. Each entry is a test
 * against the normalised header plus the label the column is given if
 * it matches — the label is not the header, because "Completion time"
 * is Forms' word for it and "Submitted" is everybody else's.
 */
const IDENTITY = [
  {
    key: 'email',
    label: 'Email',
    test: /\b(email|e mail|email address|respondent|respondent email|username|user)\b/,
  },
  {
    key: 'submitted',
    label: 'Submitted',
    test: /\b(completion time|submitted|submission time|submit time|date submitted|timestamp|date answered|answered on|answered at|response date|end time|start time)\b/,
  },
  {
    key: 'department',
    label: 'Department',
    test: /\b(department|dept|division|unit|section|team|branch|site|location|plant)\b/,
  },
];

/**
 * Splits a dataset's columns into who-answered and what-they-answered.
 *
 * Returns `{ identity, questions }`. `identity` holds at most one
 * column per key, in IDENTITY order, each carrying the label it is
 * shown under; a sheet missing one of them simply gets a shorter list
 * rather than a blank column. `questions` is everything else in sheet
 * order — including the columns the charts parked, because a question
 * nobody charted is still a question somebody answered.
 *
 * First match wins per key. A form with both "Start time" and
 * "Completion time" has one submission time as far as a reader is
 * concerned, and showing both would push the answers off the screen to
 * say the same thing twice.
 */
export function splitResponseColumns(dataset) {
  const columns = dataset?.columns ?? [];
  const identity = [];
  const taken = new Set();

  for (const { key, label, test } of IDENTITY) {
    const found = columns.find(
      (column) => !taken.has(column.name) && test.test(normalizeHeader(column.name)),
    );
    if (!found) continue;
    taken.add(found.name);
    identity.push({ ...found, key, label });
  }

  return { identity, questions: columns.filter((c) => !taken.has(c.name)) };
}

/**
 * The columns one row of the response TABLE shows, in reading order.
 *
 * `limit` caps the QUESTIONS only. The identity columns are never
 * dropped: they are the shortest columns on the sheet and the reason
 * the table is being read at all, so a cap that could hide the email
 * address would defeat the panel it is meant to keep narrow. Pass null
 * for every question.
 */
export function responseTableColumns(dataset, limit = 5) {
  const { identity, questions } = splitResponseColumns(dataset);
  return [...identity, ...(limit === null ? questions : questions.slice(0, limit))];
}

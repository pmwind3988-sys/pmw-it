// The columns a form adds that nobody asked a question about.
//
// Every survey export arrives with a band of bookkeeping down the left:
// when it was submitted, how long it took, which address submitted it,
// a running number. None of it is an answer, all of it profiles cleanly
// as a real temporal or a real measure, and left alone it wins chart
// slots from the columns the survey was actually about -- a bar chart
// of "rows by Timestamp" is the canonical wasted tile.
//
// Matched on the HEADER, never on shape. Shape cannot tell a submission
// timestamp from an incident date; both are datetimes with one value per
// row. The name is the only place the distinction is recorded, so the
// name is what is read -- and a match is hidden, never deleted, because
// a lexicon that is wrong about one column must cost the user a click
// to undo rather than a re-import.

// Header text, lowercased with punctuation flattened to single spaces,
// so `Submitted at:` and `submitted_at` both arrive as `submitted at`.
export function normalizeHeader(name) {
  return String(name ?? '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

// Each entry is [test, reason]. The reason is shown to the user, so it
// says what the column records rather than naming the rule that caught
// it.
const RULES = [
  [
    /\b(timestamp|submitted|submission time|submit time|date submitted|date answered|answered on|answered at|response date|date of response|completion time|start time|end time|date created|created at|created on|last modified|modified on|updated at|updated on|recorded at)\b/,
    'records when the form was filled in, not what it says',
  ],
  [
    /\b(time taken|time spent|time to complete|duration|elapsed|seconds taken|minutes taken)\b/,
    'measures how long the form took to fill in',
  ],
  [
    /\b(e mail|email|email address|respondent|respondent id|response id|submission id|entry id|user id|username|ip address|record id|record locator)\b/,
    'identifies who answered, which is not part of the answer',
  ],
  [
    // Anchored, unlike the rule above. A bare `\bname\b` would also
    // catch "Name of the machine that failed", which is an answer.
    /^(name|your name|full name|first name|last name|surname|respondent name)$/,
    'is who answered, which is not part of the answer',
  ],
  [
    /^(no|nos|bil|num|number|sn|s n|seq|sequence|index|id|row|row number|item no)$/,
    'a running number, not a measurement',
  ],
  [
    /^(points|total points|score out of|quiz feedback)$/,
    'form scoring added by the tool, not by a respondent',
  ],
  [
    /^(unnamed( \d+)?|column\s*\d*|field\s*\d*|)$/,
    'has no header, so nothing on the canvas could label it',
  ],
];

/**
 * The bookkeeping columns in a profile, with the reason each was picked.
 *
 * Columns already sitting at role `ignored` are left out: the profiler
 * has parked them, hiding them again would say nothing, and listing them
 * in the brief would pad "4 columns hidden" with columns the user was
 * never going to see anyway.
 */
export function detectAdminColumns(profile) {
  const found = [];

  for (const column of profile?.columns ?? []) {
    if (column.role === 'ignored') continue;

    const header = normalizeHeader(column.name);
    for (const [test, reason] of RULES) {
      if (!test.test(header)) continue;
      found.push({ name: column.name, index: column.index, reason });
      break;
    }
  }

  return found;
}

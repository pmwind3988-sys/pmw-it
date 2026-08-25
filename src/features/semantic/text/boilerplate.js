// What a survey answer says that is not an answer -- spec §6.2.
//
// Two rules, and the second is the one that can quietly destroy data.
// "no issue from IT" is somebody saying nothing is wrong; "No proper
// system exists for tracking approvals" is somebody reporting a real
// problem that happens to start with the same word. A prefix match
// alone deletes the second, so the rule needs BOTH a leading
// non-answer word and a short body -- a real complaint is never four
// words long.

// Escape sequences, never literals: an invisible character in source
// does not survive being retyped or diffed and rots into a no-op.
const ZERO_WIDTH_RE = /[\u200B-\u200D\uFEFF]/g;

export function normalizeText(value) {
  return String(value ?? '')
    .normalize('NFKC')
    .replace(ZERO_WIDTH_RE, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// Respondents paste the question into the answer box. The bracket is
// optional on both sides because the real export has rows missing one.
const LABEL_RE = /^\s*\[?\s*(selected\s+challenge|detailed\s+description|challenge|description|issue|problem)s?\s*\]?\s*:\s*/i;

export function stripLabelPrefix(line) {
  return normalizeText(line).replace(LABEL_RE, '').trim();
}

const NON_ANSWER_WORDS = new Set([
  'no', 'none', 'nil', 'na', 'nothing', 'nope', 'n/a', '-', '–', '—', '.',
]);

// Short enough that it cannot be a report of anything. Twenty letters is
// "no issue from IT" (13) with room to spare, and well under the
// shortest real complaint in the source data.
const NON_ANSWER_MAX_LETTERS = 20;

export function isNonAnswer(fragment) {
  const text = normalizeText(fragment);
  if (text === '') return true;

  const letters = text.replace(/[^A-Za-z]/g, '').length;
  if (letters === 0) return true;

  const first = text.toLowerCase().split(/[\s,.;:!?]+/)[0];
  return NON_ANSWER_WORDS.has(first) && letters <= NON_ANSWER_MAX_LETTERS;
}

/**
 * Every email sent (or tried) for one shared checklist, kept on the link row
 * in `EmailLog` as JSON, newest last.
 *
 * An entry: `{ at, kind, to, by, ok, error }`
 *   kind  'signed'  sent by the server the moment the employee signed
 *         'resent'  sent again by somebody in IT (`by`)
 *   ok    false when the send failed; `error` says why, in a line
 *
 * A checklist signed before emails existed has no log, which is exactly how
 * the panel tells "never emailed" from "emailed".
 *
 * Pure; the server appends, the portal reads.
 */

export const EMAIL_KINDS = { SIGNED: 'signed', RESENT: 'resent' };

// The log rides on the link row; a link resent every day for a year must not
// grow it without end. The oldest entries go first.
export const MAX_LOG_ENTRIES = 20;

const isEntry = (entry) => entry && typeof entry === 'object' && typeof entry.at === 'string';

export function parseEmailLog(value) {
  if (Array.isArray(value)) return value.filter(isEntry);
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed.filter(isEntry) : [];
  } catch {
    return [];
  }
}

export function appendEmailLog(log, entry) {
  const clean = {
    at: entry.at,
    kind: entry.kind,
    to: Array.isArray(entry.to) ? entry.to : [],
    by: entry.by ?? '',
    ok: entry.ok !== false,
    ...(entry.ok === false ? { error: String(entry.error ?? 'Not sent').slice(0, 300) } : {}),
  };
  return [...parseEmailLog(log), clean].slice(-MAX_LOG_ENTRIES);
}

export const serialiseEmailLog = (log) => JSON.stringify(parseEmailLog(log));

/**
 * What the panel shows: `sent` when the newest attempt went out, `failed`
 * when the newest attempt did not (a later success clears an earlier
 * failure), `none` when nothing was ever tried.
 */
export function emailStatus(linkOrLog) {
  const log = parseEmailLog(Array.isArray(linkOrLog) || typeof linkOrLog === 'string'
    ? linkOrLog
    : linkOrLog?.emailLog);
  if (!log.length) return { state: 'none', last: null, lastSent: null, log };
  const last = log[log.length - 1];
  const lastSent = [...log].reverse().find((entry) => entry.ok) ?? null;
  return { state: last.ok ? 'sent' : 'failed', last, lastSent, log };
}

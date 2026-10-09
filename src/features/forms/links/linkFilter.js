import { linkState } from './linkRules.js';
import { OUT, modeLabel } from '../checklistForm.js';

/**
 * What the shared-checklists page shows, and how its timeline is arranged,
 * as pure functions over plain link objects.
 *
 * The query string carries the choices, so a link to a filtered view survives
 * a reload and the till can point straight at its own checklists:
 *   show  all | waiting | signed | closed   (closed = expired or cancelled)
 *   from  any | till | share               (till = made for a handover or return)
 *   kind  any | handout | return           (return = a form in the OUT mode)
 *   q     free text over name, form type, items and serials
 */

export const SHOW = ['all', 'waiting', 'signed', 'closed'];
export const SOURCES = ['any', 'till', 'share'];
export const KINDS = ['any', 'handout', 'return'];

export const fromTill = (link) => Array.isArray(link?.handovers?.ids) && link.handovers.ids.length > 0;

const STATES = {
  waiting: new Set(['open', 'busy']),
  signed: new Set(['signed']),
  closed: new Set(['expired', 'cancelled']),
};

const DAY = 86400000;
const SERIAL_SEPARATOR = ' · S/N ';

export function readFilter(params) {
  const show = params.get('show');
  const from = params.get('from');
  const kind = params.get('kind');
  return {
    show: SHOW.includes(show) ? show : 'all',
    source: SOURCES.includes(from) ? from : 'any',
    kind: KINDS.includes(kind) ? kind : 'any',
    q: params.get('q') ?? '',
  };
}

/** What a link covered: the values the employee signed, else what IT pre-filled. */
const valuesOf = (link) => link?.submitted ?? link?.preset ?? {};

/** The items a link covers, as name and quantity. */
export function linkItems(link) {
  const values = valuesOf(link);
  if (Array.isArray(values.items)) {
    const named = values.items
      .filter((row) => String(row?.item ?? '').trim())
      .map((row) => ({ name: String(row.item).trim(), qty: Number(row.quantity) || 1 }));
    if (named.length) return named;
  }
  if (Array.isArray(values.checkedItems)) {
    return values.checkedItems
      .filter((name) => String(name ?? '').trim())
      .map((name) => ({ name: String(name).trim(), qty: 1 }));
  }
  return [];
}

/** Serial lines, each split into what it is and its serial where it says so. */
export function linkSerials(link) {
  return String(valuesOf(link).serialNumbers ?? '')
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const at = line.indexOf(SERIAL_SEPARATOR);
      if (at < 0) return { what: '', serial: line };
      return {
        what: line.slice(0, at).trim(),
        serial: line.slice(at + SERIAL_SEPARATOR.length).trim(),
      };
    });
}

function matchesQuery(link, query) {
  const needle = query.trim().toLowerCase();
  if (!needle) return true;
  const haystack = [
    link.employeeName,
    modeLabel(link.formMode),
    ...linkItems(link).map((item) => item.name),
    ...linkSerials(link).map((line) => `${line.what} ${line.serial}`),
  ].join('\n').toLowerCase();
  return haystack.includes(needle);
}

/** Applies show, source, kind and the text query. `till` is the older boolean switch. */
export function filterLinks(links, filter = {}, now = Date.now()) {
  const {
    show = 'all',
    kind = 'any',
    q = '',
  } = filter;
  const source = filter.source ?? (filter.till ? 'till' : 'any');

  return (links ?? []).filter((link) => {
    if (source === 'till' && !fromTill(link)) return false;
    if (source === 'share' && fromTill(link)) return false;
    if (kind === 'return' && link.formMode !== OUT) return false;
    if (kind === 'handout' && link.formMode === OUT) return false;
    if (!matchesQuery(link, q)) return false;
    if (show === 'all') return true;
    return STATES[show].has(linkState(link, now));
  });
}

/** How many each status bubble would show, under the other filters as they stand. */
export function countLinks(links, filter = {}, now = Date.now()) {
  return Object.fromEntries(
    SHOW.map((show) => [show, filterLinks(links, { ...filter, show }, now).length]),
  );
}

/** The timeline's label for a moment: Today, Yesterday, ..., or a month. */
export function dayGroupLabel(time, now = Date.now()) {
  const at = typeof time === 'number' ? time : Date.parse(time);
  if (!Number.isFinite(at)) return 'Undated';

  const when = new Date(at);
  const today = new Date(now);
  const startOf = (date) => new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
  const days = Math.round((startOf(today) - startOf(when)) / DAY);

  if (days <= 0) return 'Today';
  if (days === 1) return 'Yesterday';
  if (days < 7) return 'Earlier this week';
  if (when.getMonth() === today.getMonth() && when.getFullYear() === today.getFullYear()) {
    return 'Earlier this month';
  }
  return when.toLocaleDateString('en-MY', { month: 'long', year: 'numeric' });
}

/** The moment a link is listed under: when it was signed, else when it was sent. */
export function linkTime(link) {
  if (linkState(link) === 'signed') {
    const signed = Date.parse(link.signedOn);
    if (Number.isFinite(signed)) return signed;
  }
  return Date.parse(link.created);
}

/** Consecutive links under the same day label, in the order given (newest first). */
export function groupByDay(links, now = Date.now()) {
  // Grouped by when each link last moved (signed, else sent), so the list is
  // re-sorted on that same time: sorted by creation, a link sent last week
  // and signed today would open "Earlier this week" above "Today".
  const time = (link) => {
    const value = linkTime(link);
    return Number.isFinite(value) ? value : -Infinity;
  };
  const sorted = [...(links ?? [])].sort((a, b) => time(b) - time(a));
  const groups = new Map();
  for (const link of sorted) {
    const label = dayGroupLabel(linkTime(link), now);
    if (!groups.has(label)) groups.set(label, []);
    groups.get(label).push(link);
  }
  return [...groups].map(([label, list]) => ({ label, links: list }));
}

const finiteOrNull = (value) => {
  const time = Date.parse(value);
  return Number.isFinite(time) ? time : null;
};

/** The story of one link, from sent to where it stands now. */
export function linkSteps(link, now = Date.now()) {
  const sent = finiteOrNull(link.created);
  const steps = [{ label: 'Sent', when: sent, done: true }];
  if (fromTill(link)) steps.push({ label: 'Handed over', when: sent, done: true });

  const state = linkState(link, now);
  if (state === 'signed') {
    steps.push({ label: 'Signed', when: finiteOrNull(link.signedOn), done: true });
  } else if (state === 'expired') {
    steps.push({ label: 'Expired', when: finiteOrNull(link.expiresOn), done: false });
  } else if (state === 'cancelled') {
    steps.push({ label: 'Cancelled', when: null, done: false });
  } else {
    steps.push({ label: 'Signed', when: null, done: false });
  }
  return steps;
}

/** One of six avatar colours, the same for the same name every time. */
export function avatarTone(name) {
  let hash = 0;
  for (const char of String(name ?? '')) {
    hash = (hash * 31 + char.codePointAt(0)) >>> 0;
  }
  return hash % 6;
}

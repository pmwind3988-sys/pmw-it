import { isBackgroundComponent } from './derive/deriveSoftware.js';
import { dayLabel } from './dayLabel.js';

/**
 * The installed-software panel's rules, pure: which programs are runtimes and
 * drivers, what the search keeps, how a match is marked, and what "Copy list"
 * puts on the clipboard.
 */

const DRIVER_WORDS = /driver|chipset|graphics|audio|management engine|software package|install/i;

/** `driver` or `runtime` beside a folded program; null for one the person chose. */
export function kindTagOf(entry) {
  if (!isBackgroundComponent(entry?.name, entry?.publisher)) return null;
  return DRIVER_WORDS.test(String(entry?.name ?? '')) ? 'driver' : 'runtime';
}

/** How many programs the "Hide runtimes and drivers (n)" switch would fold away. */
export function backgroundCount(list) {
  return (list ?? []).filter((entry) => kindTagOf(entry)).length;
}

/**
 * The rows the panel shows: the search (name or publisher, any case), and the
 * runtimes folded away when the switch is on. Order is the scan's own.
 */
export function visibleSoftware(list, { query = '', hideRuntimes = true } = {}) {
  const needle = String(query ?? '').trim().toLowerCase();
  return (list ?? []).filter((entry) => {
    if (hideRuntimes && kindTagOf(entry)) return false;
    if (!needle) return true;
    return `${entry.name ?? ''} ${entry.publisher ?? ''}`.toLowerCase().includes(needle);
  });
}

/**
 * The name split around the search text, so the page can bold each match.
 * Returns [{ text, match }] covering the whole string, in order.
 */
export function highlightParts(text, query) {
  const source = String(text ?? '');
  const needle = String(query ?? '').trim();
  if (!needle) return [{ text: source, match: false }];

  const lower = source.toLowerCase();
  const target = needle.toLowerCase();
  const parts = [];
  let at = 0;
  let hit = lower.indexOf(target);
  while (hit !== -1) {
    if (hit > at) parts.push({ text: source.slice(at, hit), match: false });
    parts.push({ text: source.slice(hit, hit + target.length), match: true });
    at = hit + target.length;
    hit = lower.indexOf(target, at);
  }
  if (at < source.length) parts.push({ text: source.slice(at), match: false });
  return parts.length ? parts : [{ text: source, match: false }];
}

/** One program per line, `Name version`, as pasted into an email or Excel. */
export function copyTextOf(rows) {
  return (rows ?? [])
    .map((entry) => [entry.name, entry.version].filter(Boolean).join(' '))
    .join('\n');
}

/** The Installed column: `2 Sep 2026`, or a dash when the scan left it blank. */
export function installedLabel(iso) {
  return dayLabel(iso) ?? '—';
}

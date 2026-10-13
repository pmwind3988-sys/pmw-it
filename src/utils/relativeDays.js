const DAY_MS = 86400000;

const startOfDay = (ms) => {
  const d = new Date(ms);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
};

/**
 * How long ago a moment was, in the words a person would use: "today",
 * "yesterday", "9 days ago", "4 months ago". Counts calendar days, so a scan
 * at 11pm yesterday reads "yesterday" at 8am today. Empty for anything that
 * is not a real instant, so callers never print "NaN".
 */
export function relativeDays(ms, now = Date.now()) {
  if (typeof ms !== 'number' || !Number.isFinite(ms)) return '';
  const days = Math.round((startOfDay(now) - startOfDay(ms)) / DAY_MS);
  if (days <= 0) return 'today';
  if (days === 1) return 'yesterday';
  if (days < 60) return `${days} days ago`;
  return `${Math.floor(days / 30.4)} months ago`;
}

/**
 * Per-browser memory for the till. Every read and write is guarded: in a
 * private window storage can throw, and the till must still work — it just
 * forgets between visits.
 */
export function readPref(key) {
  try { return localStorage.getItem(key); } catch { return null; }
}

export function writePref(key, value) {
  try {
    if (value == null) localStorage.removeItem(key);
    else localStorage.setItem(key, value);
  } catch { /* private mode */ }
}

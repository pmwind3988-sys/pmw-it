/**
 * Which build of the portal this tab is running, and whether the server now
 * has a newer one.
 *
 * Every deploy writes `/version.json` (see `appVersion` in vite.config.js)
 * holding the same id that was baked into the bundle. An installed app can
 * stay open for days; comparing the two is how it finds out it is behind.
 */

/* global __APP_VERSION__ */

/** Baked in at build time. `dev` under `npm run dev`, where nothing is checked. */
export const BUILT_VERSION = typeof __APP_VERSION__ === 'string' ? __APP_VERSION__ : 'dev';

/**
 * True only when the server answered with a real id that differs from ours.
 * A failed fetch, a missing file or an SPA fallback page (HTML, no id) is NOT
 * a new version -- reloading on those would loop on a bad network.
 */
export function isNewerBuild(built, served) {
  if (!built || built === 'dev') return false;
  if (typeof served !== 'string' || !served.trim()) return false;
  return served.trim() !== built;
}

/**
 * A launch that finds itself behind reloads at once -- nothing has been typed
 * yet -- but only once per deployed id. If the reload still comes back old (a
 * CDN edge that has not caught up), the bar takes over instead of a loop.
 */
export function shouldReloadOnLaunch(built, served, reloadedFor) {
  return isNewerBuild(built, served) && reloadedFor !== served.trim();
}

/**
 * The Android app opens the site with an `android-app://<package>` referrer;
 * a home-screen install runs in `standalone` display mode. The referrer is
 * gone after the first navigation (the Microsoft sign-in round trip, say), so
 * the answer is remembered for the tab's life.
 */
export function isInstalledApp({ referrer = '', standalone = false, remembered = false } = {}) {
  return remembered || standalone || referrer.startsWith('android-app://');
}

const INSTALLED_KEY = 'pmw:installedApp';
const RELOADED_FOR_KEY = 'pmw:reloadedForVersion';

function session() {
  try {
    return window.sessionStorage;
  } catch {
    return null;
  }
}

/** Called once at start-up, while the launch referrer is still there. */
export function rememberInstalledApp() {
  const installed = isInstalledApp({
    referrer: document.referrer,
    standalone: window.matchMedia?.('(display-mode: standalone)').matches ?? false,
    remembered: session()?.getItem(INSTALLED_KEY) === '1',
  });
  try {
    if (installed) session()?.setItem(INSTALLED_KEY, '1');
  } catch {
    // Private mode: the button just goes missing after a redirect.
  }
  return installed;
}

export function runningAsInstalledApp() {
  return session()?.getItem(INSTALLED_KEY) === '1';
}

export function reloadedForVersion() {
  return session()?.getItem(RELOADED_FOR_KEY) ?? null;
}

/**
 * Notes which build a launch is reloading for. False when it cannot be noted
 * (storage blocked) -- the launch then must NOT reload, or a CDN edge still
 * serving the old build would reload it forever.
 */
export function noteReloadFor(served) {
  try {
    const store = session();
    if (!store) return false;
    store.setItem(RELOADED_FOR_KEY, served.trim());
    return true;
  } catch {
    return false;
  }
}

/** Reads the deployed id, or null when it cannot be told. Never throws. */
export async function fetchServedVersion(fetchImpl = globalThis.fetch) {
  try {
    const response = await fetchImpl(`/version.json?t=${Date.now()}`, { cache: 'no-store' });
    if (!response.ok) return null;
    const body = await response.json();
    return typeof body?.version === 'string' ? body.version : null;
  } catch {
    return null;
  }
}

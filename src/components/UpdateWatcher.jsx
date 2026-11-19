import { useEffect, useRef, useState } from 'react';
import { useLocation } from 'react-router-dom';
import {
  BUILT_VERSION,
  fetchServedVersion,
  isNewerBuild,
  noteReloadFor,
  reloadedForVersion,
  shouldReloadOnLaunch,
} from '../utils/appVersion';

const CHECK_EVERY_MS = 5 * 60 * 1000;
// A launch check that answers later than this may already have somebody
// typing, so it falls back to the bar like any other check.
const LAUNCH_WINDOW_MS = 8000;

/**
 * Keeps an installed app (the APK, or a home-screen install) on the latest
 * deploy of main.
 *
 * - On launch it checks at once, and if it is behind it reloads straight into
 *   the new build: nothing has been typed yet.
 * - Left open, it checks every five minutes and whenever it comes back to the
 *   foreground. A newer build then shows a bar, and the next move to another
 *   screen reloads into it. It never reloads under somebody mid-screen on its
 *   own, because a half-filled form or till receipt would be lost; the bar's
 *   button does it at once.
 *
 * The installed app's own "Check for updates" button lives in the nav column
 * (`UpdateCheckButton`).
 */
export default function UpdateWatcher() {
  const [stale, setStale] = useState(false);
  const { key } = useLocation();
  const firstKey = useRef(key);

  useEffect(() => {
    if (BUILT_VERSION === 'dev') return undefined;
    let cancelled = false;
    const launchedAt = Date.now();

    const check = async () => {
      if (document.hidden) return;
      const served = await fetchServedVersion();
      if (cancelled || !isNewerBuild(BUILT_VERSION, served)) return;
      const justLaunched = Date.now() - launchedAt < LAUNCH_WINDOW_MS;
      if (
        justLaunched &&
        shouldReloadOnLaunch(BUILT_VERSION, served, reloadedForVersion()) &&
        noteReloadFor(served)
      ) {
        window.location.reload();
        return;
      }
      setStale(true);
    };
    const onVisible = () => {
      if (!document.hidden) check();
    };

    const timer = window.setInterval(check, CHECK_EVERY_MS);
    document.addEventListener('visibilitychange', onVisible);
    window.addEventListener('online', check);
    check();
    return () => {
      cancelled = true;
      window.clearInterval(timer);
      document.removeEventListener('visibilitychange', onVisible);
      window.removeEventListener('online', check);
    };
  }, []);

  // Moving to another screen is the safe moment: whatever was on the old one
  // has already been left behind.
  useEffect(() => {
    if (stale && key !== firstKey.current) window.location.reload();
  }, [stale, key]);

  if (!stale) return null;
  return (
    <div className="update-bar" role="status">
      <span>A new version of the portal is ready.</span>
      <button type="button" onClick={() => window.location.reload()}>
        Update now
      </button>
    </div>
  );
}

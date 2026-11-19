import { useEffect, useState } from 'react';
import { BUILT_VERSION, fetchServedVersion, isNewerBuild, runningAsInstalledApp } from '../utils/appVersion';
import Spinner from './ui/Spinner';
import { RefreshCw } from './ui/Icons';

const MESSAGES = {
  latest: 'Up to date',
  offline: "Couldn't check",
};

/**
 * "Check for updates", in the nav column of the INSTALLED app only (the APK or
 * a home-screen install). In a browser tab the browser's own reload does the
 * job, so the button would only be noise there.
 *
 * Pressing it is somebody asking, so a newer build is loaded at once.
 */
export default function UpdateCheckButton() {
  const [installed] = useState(runningAsInstalledApp);
  const [state, setState] = useState('idle');

  useEffect(() => {
    if (state !== 'latest' && state !== 'offline') return undefined;
    const timer = window.setTimeout(() => setState('idle'), 4000);
    return () => window.clearTimeout(timer);
  }, [state]);

  if (!installed) return null;

  const check = async () => {
    setState('checking');
    const served = await fetchServedVersion();
    if (served === null) return setState('offline');
    if (isNewerBuild(BUILT_VERSION, served)) {
      window.location.reload();
      return undefined;
    }
    return setState('latest');
  };

  return (
    <button
      type="button"
      className="shell-signout shell-updatecheck"
      onClick={check}
      disabled={state === 'checking'}
      aria-live="polite"
    >
      {state === 'checking' ? <Spinner size={13} /> : <RefreshCw size={13} />}
      {MESSAGES[state] ?? (state === 'checking' ? 'Checking…' : 'Check for updates')}
    </button>
  );
}

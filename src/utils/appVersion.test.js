import { describe, expect, it } from 'vitest';
import { fetchServedVersion, isInstalledApp, isNewerBuild, shouldReloadOnLaunch } from './appVersion.js';

describe('isNewerBuild', () => {
  it('is true when the server has a different build', () => {
    expect(isNewerBuild('abc123', 'def456')).toBe(true);
  });

  it('is false for the same build, ignoring stray whitespace', () => {
    expect(isNewerBuild('abc123', 'abc123')).toBe(false);
    expect(isNewerBuild('abc123', ' abc123\n')).toBe(false);
  });

  it('never fires in dev or when the server said nothing usable', () => {
    expect(isNewerBuild('dev', 'def456')).toBe(false);
    expect(isNewerBuild('abc123', null)).toBe(false);
    expect(isNewerBuild('abc123', '')).toBe(false);
    expect(isNewerBuild('abc123', undefined)).toBe(false);
  });
});

describe('fetchServedVersion', () => {
  const answer = (ok, body) => async () => ({
    ok,
    json: async () => {
      if (body instanceof Error) throw body;
      return body;
    },
  });

  it('reads the id', async () => {
    expect(await fetchServedVersion(answer(true, { version: 'def456' }))).toBe('def456');
  });

  it('is null for an error status, a non-JSON page or a network failure', async () => {
    expect(await fetchServedVersion(answer(false, { version: 'x' }))).toBe(null);
    expect(await fetchServedVersion(answer(true, new SyntaxError('<!doctype')))).toBe(null);
    expect(
      await fetchServedVersion(async () => {
        throw new TypeError('Failed to fetch');
      }),
    ).toBe(null);
  });

  it('is null when the file has no id', async () => {
    expect(await fetchServedVersion(answer(true, {}))).toBe(null);
  });
});

describe('shouldReloadOnLaunch', () => {
  it('reloads a launch that is behind', () => {
    expect(shouldReloadOnLaunch('abc123', 'def456', null)).toBe(true);
  });

  it('reloads only once per deployed build, so a lagging CDN cannot loop it', () => {
    expect(shouldReloadOnLaunch('abc123', 'def456', 'def456')).toBe(false);
    expect(shouldReloadOnLaunch('abc123', 'ghi789', 'def456')).toBe(true);
  });

  it('never reloads a launch that is current or could not check', () => {
    expect(shouldReloadOnLaunch('abc123', 'abc123', null)).toBe(false);
    expect(shouldReloadOnLaunch('abc123', null, null)).toBe(false);
    expect(shouldReloadOnLaunch('dev', 'def456', null)).toBe(false);
  });
});

describe('isInstalledApp', () => {
  it('knows the Android app by its launch referrer', () => {
    expect(isInstalledApp({ referrer: 'android-app://com.pmwgroup.itportal/' })).toBe(true);
  });

  it('knows a home-screen install, and remembers past a sign-in redirect', () => {
    expect(isInstalledApp({ standalone: true })).toBe(true);
    expect(isInstalledApp({ referrer: 'https://login.microsoftonline.com/', remembered: true })).toBe(true);
  });

  it('is false in an ordinary browser tab', () => {
    expect(isInstalledApp({ referrer: 'https://pmw-it.vercel.app/dashboard' })).toBe(false);
    expect(isInstalledApp()).toBe(false);
  });
});

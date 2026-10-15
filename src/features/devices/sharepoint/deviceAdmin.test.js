import { describe, it, expect, afterEach, vi } from 'vitest';
import { isDeviceAdmin } from './deviceAdmin.js';

const reply = (body, status = 200) => ({ ok: status < 300, status, json: async () => body, text: async () => 'x', headers: { get: () => null } });
const perms = (low, high = 0) => ({ d: { EffectiveBasePermissions: { Low: String(low), High: String(high) } } });

afterEach(() => vi.unstubAllGlobals());

describe('isDeviceAdmin', () => {
  it('is yes for a site owner (Full Control carries ManagePermissions)', async () => {
    vi.stubGlobal('fetch', async () => reply(perms(1 << 25)));
    expect(await isDeviceAdmin('https://x', 't')).toBe(true);
  });

  it('is no for a member who can edit and delete list items', async () => {
    // Add, edit, delete items and manage lists -- everything Edit gives, nothing more.
    vi.stubGlobal('fetch', async () => reply(perms(0b1111 | (1 << 11))));
    expect(await isDeviceAdmin('https://x', 't')).toBe(false);
  });

  it('is no when the permissions cannot be read', async () => {
    vi.stubGlobal('fetch', async () => reply({}, 403));
    expect(await isDeviceAdmin('https://x', 't')).toBe(false);
  });
});

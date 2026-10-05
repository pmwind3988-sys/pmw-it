import { describe, it, expect, afterEach, vi } from 'vitest';
import { readAssignments } from './readAssignments.js';
import { readDevice } from './readDevices.js';

const reply = (body, status = 200) => ({ ok: status < 300, status, json: async () => body });

describe('readAssignments', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('reads every page', async () => {
    const urls = [];
    vi.stubGlobal('fetch', async (url) => {
      urls.push(url);
      return urls.length === 1
        ? reply({ d: { results: [{ Id: 1, DeviceId: 4, Owner: 'Amir' }], __next: 'https://x/next' } })
        : reply({ d: { results: [{ Id: 2, DeviceId: 4, Owner: 'Farid' }] } });
    });
    const stints = await readAssignments('https://x', 't');
    expect(stints.map((s) => s.owner)).toEqual(['Amir', 'Farid']);
  });

  it('asks for one machine when given its id', async () => {
    let asked = '';
    vi.stubGlobal('fetch', async (url) => { asked = url; return reply({ d: { results: [] } }); });
    await readAssignments('https://x', 't', { deviceId: 9 });
    expect(decodeURIComponent(asked)).toContain('$filter=DeviceId eq 9');
  });

  it('reads a list that does not exist yet as no history', async () => {
    vi.stubGlobal('fetch', async () => reply({}, 404));
    await expect(readAssignments('https://x', 't')).resolves.toEqual([]);
  });
});

describe('readDevice', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('reads one row by id', async () => {
    vi.stubGlobal('fetch', async () => reply({ d: { Id: 7, Title: 'PC1', Status: 'Spare' } }));
    const device = await readDevice('https://x', 't', 7);
    expect(device).toMatchObject({ id: 7, computerName: 'PC1', status: 'Spare' });
  });

  it('answers null for a row that has gone', async () => {
    vi.stubGlobal('fetch', async () => reply({}, 404));
    await expect(readDevice('https://x', 't', 7)).resolves.toBeNull();
  });
});

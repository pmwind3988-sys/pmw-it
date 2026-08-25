import { describe, it, expect } from 'vitest';
import { serverDependency } from './serverDependency.js';

describe('serverDependency', () => {
  it('reads mapped drives as working off the server', () => {
    const result = serverDependency({ mappedDrives: 3, networkType: 'Ethernet' });
    expect(result.serverDependent).toBe(true);
    expect(result.networkRisk).toBe('Fine');
  });

  it('calls server work over Wi-Fi from a desktop a severe bottleneck', () => {
    const result = serverDependency({ mappedDrives: 2, networkType: 'Wi-Fi', deviceType: 'Desktop' });
    expect(result.networkRisk).toBe('Severe');
    expect(result.networkNote).toMatch(/cable/);
  });

  it('expects a laptop to be on Wi-Fi rather than calling it a fault', () => {
    const result = serverDependency({ mappedDrives: 2, networkType: 'Wi-Fi', deviceType: 'Laptop' });
    expect(result.networkRisk).toBe('Wireless');
  });

  it('counts saved server credentials even with no drive mapped', () => {
    const result = serverDependency({ mappedDrives: 0, serverCredentials: 'PMWSERVER | amir' });
    expect(result.serverDependent).toBe(true);
  });

  it('leaves a standalone machine out of the network question', () => {
    const result = serverDependency({ mappedDrives: 0, networkType: 'Wi-Fi' });
    expect(result.serverDependent).toBe(false);
    expect(result.networkRisk).toBe('None');
  });

  it('does not guess the link when the scan did not report one', () => {
    expect(serverDependency({ mappedDrives: 1 }).networkRisk).toBe('Unknown');
  });
});

import { describe, it, expect } from 'vitest';
import { gpuClass } from './gpuClass.js';

describe('gpuClass', () => {
  it('finds the card behind an integrated adapter listed first', () => {
    const result = gpuClass(['Intel(R) UHD Graphics', 'NVIDIA GeForce RTX 3060']);
    expect(result.gpuClass).toBe('Dedicated');
    expect(result.dedicatedGpuName).toMatch(/RTX 3060/);
  });

  it('calls processor graphics integrated', () => {
    expect(gpuClass(['Intel(R) UHD Graphics']).dedicatedGpu).toBe(false);
    expect(gpuClass(['AMD Radeon(TM) Graphics']).gpuClass).toBe('Integrated');
  });

  it('does not guess when nothing was reported', () => {
    expect(gpuClass([]).dedicatedGpu).toBe(null);
    expect(gpuClass(['Some Unknown Adapter']).gpuClass).toBe('Unknown');
  });
});

import { describe, it, expect } from 'vitest';
import { brandOf, softwareBrandOf, initialsOfBrand, logoUrlOf, partBrandOf } from './brandLogo.js';

describe('brandOf', () => {
  it('reads the maker out of the text the scan prints', () => {
    expect(brandOf('Intel(R) Core(TM) i5-1135G7').slug).toBe('intel');
    expect(brandOf('AMD Ryzen 5 7530U').slug).toBe('amd');
    expect(brandOf('NVIDIA GeForce RTX 3050').slug).toBe('nvidia');
    expect(brandOf('ASUSTeK COMPUTER INC.').slug).toBe('asus');
    expect(brandOf('WDC WD10JPVX-60JC3T1').slug).toBe('westerndigital');
    expect(brandOf('Samsung SSD 870 EVO').slug).toBe('samsung');
  });
  it('knows the memory maker codes', () => {
    expect(brandOf('80AD').name).toBe('SK hynix');
  });
  it('does not take "hard" or other words for a brand', () => {
    expect(brandOf('Generic hardware thing')).toBeNull();
    expect(brandOf('', null)).toBeNull();
  });
  it('picks the specific Microsoft product before Microsoft itself', () => {
    expect(brandOf('Microsoft Teams').slug).toBe('microsoftteams');
    expect(brandOf('Microsoft 365 Apps for business').slug).toBe('microsoftoffice');
  });
});

describe('softwareBrandOf', () => {
  it('guesses a slug from the publisher when the program is not known, and says so', () => {
    expect(softwareBrandOf('Foo Tool', 'Zebra Technologies')).toEqual({ slug: 'zebra', name: 'Zebra', guessed: true });
  });
  it('guesses nothing from a word too short to mean anything', () => {
    expect(softwareBrandOf('X', '')).toBeNull();
  });
});

describe('the rest', () => {
  it('makes initials and the logo address', () => {
    expect(initialsOfBrand('Zebra Technologies')).toBe('ZT');
    expect(initialsOfBrand('')).toBe('?');
    expect(logoUrlOf('intel')).toBe('https://cdn.simpleicons.org/intel');
    expect(logoUrlOf(null)).toBeNull();
    expect(logoUrlOf('windows')).toBeNull();
  });
  it('gives Windows its own logo and RAM its sticks’ maker', () => {
    expect(partBrandOf('windows', {}, 'Windows 11 Pro').slug).toBe('windows');
    expect(partBrandOf('ram', { memoryModules: [{ maker: 'Kingston' }] }, '16 GB').slug).toBe('kingstontechnology');
  });
});

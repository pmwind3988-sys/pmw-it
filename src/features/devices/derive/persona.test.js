import { describe, it, expect } from 'vitest';
import { PERSONAS } from './persona.js';

describe('PERSONAS', () => {
  it('has the right label for the heavy profile', () => {
    expect(PERSONAS.HEAVY.label).toBe('Engineering / Technical / Media');
  });
});

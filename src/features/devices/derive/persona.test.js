import { describe, it, expect } from 'vitest';
import { personaFor, PERSONAS } from './persona.js';

describe('personaFor', () => {
  it('reads drawing departments as the heavy profile', () => {
    expect(personaFor('ENGINEERING').key).toBe('heavy');
    expect(personaFor('qaqc').key).toBe('heavy');
  });

  it('reads the people who work away from a desk as mobile', () => {
    expect(personaFor('SALES')).toBe(PERSONAS.MOBILE);
  });

  it('falls back to the desk baseline for a department it does not know', () => {
    expect(personaFor('CANTEEN')).toBe(PERSONAS.DESK);
  });

  it('marks a machine with no department as unclassified', () => {
    expect(personaFor(null)).toBe(PERSONAS.UNKNOWN);
    expect(personaFor('')).toBe(PERSONAS.UNKNOWN);
  });
});

import { describe, it, expect } from 'vitest';
import { withPreviousName } from './previousNames.js';

describe('withPreviousName', () => {
  it('adds the outgoing name after the ones already kept', () => {
    expect(withPreviousName(['OLD-1'], 'PC1', 'PC9')).toEqual(['OLD-1', 'PC1']);
    expect(withPreviousName(null, ' PC1 ', 'PC9')).toEqual(['PC1']);
  });

  it('keeps one copy of a name, whatever its case, at its latest place', () => {
    expect(withPreviousName(['pc1', 'OLD-1'], 'PC1', 'PC9')).toEqual(['OLD-1', 'PC1']);
  });

  it('takes a name off the list when the machine is given it back', () => {
    expect(withPreviousName(['OLD-1', 'PC1'], 'PC9', 'pc1')).toEqual(['OLD-1', 'PC9']);
  });

  it('adds nothing for a blank old name or a case-only change', () => {
    expect(withPreviousName([], '', 'PC9')).toEqual([]);
    expect(withPreviousName([], 'pc1', 'PC1')).toEqual([]);
  });
});

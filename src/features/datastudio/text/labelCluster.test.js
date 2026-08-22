import { describe, it, expect } from 'vitest';
import { labelCluster } from './labelCluster.js';

describe('labelCluster', () => {
  const corpus = [
    'approval requests wait for days with no reminder',
    'chasing approval status by email every week',
    'approval sign-off has no reminder or tracking',
    'reports are rebuilt from scratch every month',
    'consolidating spreadsheets from five subsidiaries',
    'the monthly report takes three days to build',
  ];

  it('names a theme after what makes it different', () => {
    const name = labelCluster(corpus.slice(0, 3), corpus);
    expect(name).toContain('approval');
    expect(name).not.toContain('report');
  });

  it('ignores a word that is in every fragment', () => {
    const flat = ['data is slow', 'data is missing', 'data is wrong'];
    // "data" appears everywhere, so it distinguishes nothing.
    const name = labelCluster(flat.slice(0, 2), flat);
    expect(name.split(' · ')).not.toContain('data');
  });

  it('drops stopwords and short noise', () => {
    const name = labelCluster(corpus.slice(0, 3), corpus);
    for (const stop of ['the', 'for', 'with', 'and', 'has', 'no']) {
      expect(name.split(' · ')).not.toContain(stop);
    }
  });

  it('says so when there is nothing distinctive to say', () => {
    expect(labelCluster([], corpus)).toBe('Unnamed theme');
  });
});

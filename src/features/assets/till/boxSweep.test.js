import { describe, it, expect } from 'vitest';
import {
  makeAndModel, readLine, newSweep, recordSweep, sweepSuggestions, sweepPending, detailLine,
} from './boxSweep.js';

describe('makeAndModel', () => {
  it('reads the model after a known make, up to the generic words', () => {
    expect(makeAndModel('Logitech M90 Wireless Mouse')).toEqual({ make: 'Logitech', model: 'M90' });
    expect(makeAndModel('Dell P2422H 24" Monitor')).toEqual({ make: 'Dell', model: 'P2422H' });
    expect(makeAndModel('Logitech Advanced Optical Tracking')).toEqual({ make: 'Logitech', model: '' });
    expect(makeAndModel('Made in China')).toBeNull();
  });
});

describe('readLine', () => {
  it('finds what kind of thing it is, its colour and its details', () => {
    const votes = Object.fromEntries(readLine('Wireless Mouse · Black · USB-C receiver'));
    expect(votes).toMatchObject({ category: 'Mouse', colour: 'Black', connection: 'Wireless', connector: 'USB-C' });
    expect(Object.fromEntries(readLine('HDMI cable 1.8 m'))).toMatchObject({ category: 'Cable', connector: 'HDMI', length: '1.8 m' });
  });
});

describe('a sweep across several sides', () => {
  const front = ['Logitech M90', 'Wired Mouse', 'Black'];
  const side = ['Logitech M90 Mouse', 'USB plug-and-play', 'Black'];
  const back = ['Logitech', 'Model: M90', 'Made in China'];

  it('offers only what was read in two passes, the most-seen value winning', () => {
    let sweep = newSweep();
    sweep = recordSweep(sweep, front);
    expect(sweepSuggestions(sweep).model).toBe('');
    expect(sweepPending(sweep).map((p) => p.kind)).toContain('model');

    sweep = recordSweep(sweep, side);
    sweep = recordSweep(sweep, back);
    expect(sweepSuggestions(sweep)).toMatchObject({
      category: 'Mouse', make: 'Logitech', model: 'M90', colour: 'Black',
    });
  });

  it('does not let one side printed five times outvote the rest', () => {
    const sweep = recordSweep(newSweep(), ['Pink', 'Pink', 'Pink', 'Pink', 'Pink']);
    expect(sweep.votes.colour).toEqual({ Pink: 1 });
  });

  it('writes the colour and details as one line', () => {
    expect(detailLine({ colour: 'Black', details: ['Wired', 'USB'] })).toBe('Black · Wired · USB');
  });
});

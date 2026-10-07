import { describe, it, expect } from 'vitest';
import { createReadGate, guessKind } from './readGate.js';

const C = (...values) => values.map((rawValue) => ({ rawValue }));

describe('createReadGate', () => {
  it('believes a code only on its second read', () => {
    const feed = createReadGate();
    expect(feed(C('5CG4211XQ7'), { aimed: true, now: 0 })).toEqual({ accept: [], choices: [] });
    expect(feed(C('5CG4211XQ7'), { aimed: false, now: 100 }).accept).toEqual(['5CG4211XQ7']);
  });

  it('ignores a one-frame misread', () => {
    const feed = createReadGate();
    feed(C('5CG4211XQ7'), { aimed: true, now: 0 });
    feed(C('5CG4Z11XQ7'), { aimed: true, now: 100 });
    // The real code reads again; the misread never does, and expires.
    expect(feed(C('5CG4211XQ7'), { aimed: true, now: 200 }).accept).toEqual([]);
    expect(feed(C('5CG4211XQ7'), { aimed: true, now: 1800 }).accept).toEqual([]);
    const later = createReadGate();
    later(C('ABC123XYZ9'), { aimed: true, now: 0 });
    later(C('ABC1Z3XYZ9'), { aimed: true, now: 50 });
    expect(later(C('ABC123XYZ9'), { aimed: true, now: 1600 }).accept).toEqual([]);
    expect(later(C('ABC123XYZ9'), { aimed: true, now: 1700 }).accept).toEqual(['ABC123XYZ9']);
  });

  it('lets the code in the aiming box win over one elsewhere in the picture', () => {
    const feed = createReadGate();
    feed(C('SHIPPING123456'), { aimed: false, now: 0 });
    feed(C('5CG4211XQ7'), { aimed: true, now: 50 });
    expect(feed(C('SHIPPING123456', '5CG4211XQ7'), { aimed: false, now: 100 }))
      .toEqual({ accept: ['5CG4211XQ7'], choices: [] });
    // The shipping label, still in view, never gets its turn while aiming.
    expect(feed(C('SHIPPING123456', '5CG4211XQ7'), { aimed: false, now: 150 })).toEqual({ accept: [], choices: [] });
    expect(feed(C('5CG4211XQ7'), { aimed: true, now: 200 })).toEqual({ accept: [], choices: [] });
  });

  it('asks when two codes are both in the aiming box', () => {
    const feed = createReadGate();
    feed(C('5CG4211XQ7', '0X7K3M'), { aimed: true, now: 0 });
    const out = feed(C('5CG4211XQ7', '0X7K3M'), { aimed: true, now: 100 });
    expect(out.accept).toEqual([]);
    expect(out.choices.sort()).toEqual(['0X7K3M', '5CG4211XQ7']);
    // Asked once, not on every frame after.
    expect(feed(C('5CG4211XQ7', '0X7K3M'), { aimed: true, now: 200 })).toEqual({ accept: [], choices: [] });
  });

  it('takes the serial without asking when a serial is what the till is waiting for', () => {
    const feed = createReadGate();
    feed(C('5CG4211XQ7', '5099206092372'), { aimed: true, now: 0 });
    expect(feed(C('5CG4211XQ7', '5099206092372'), { aimed: true, now: 100, expect: 'serial' }).accept).toEqual(['5CG4211XQ7']);
  });

  it('takes a box held still once, and again only after it was taken away', () => {
    const feed = createReadGate({ gapMs: 1000 });
    feed(C('ABC123XYZ9'), { aimed: true, now: 0 });
    expect(feed(C('ABC123XYZ9'), { aimed: true, now: 100 }).accept).toHaveLength(1);
    for (let t = 200; t <= 3000; t += 100) {
      expect(feed(C('ABC123XYZ9'), { aimed: true, now: t }).accept).toHaveLength(0);
    }
    feed([], { now: 4500 });
    feed(C('ABC123XYZ9'), { aimed: true, now: 4600 });
    expect(feed(C('ABC123XYZ9'), { aimed: true, now: 4700 }).accept).toHaveLength(1);
  });
});

describe('guessKind', () => {
  it('says what a code probably is', () => {
    expect(guessKind('5099206092372')).toMatch(/Shop barcode/);
    expect(guessKind('5CG4211XQ7')).toMatch(/serial/);
    expect(guessKind('00:1A:2B:3C:4D:5E')).toMatch(/MAC/);
  });
});

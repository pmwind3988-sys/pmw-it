import { describe, it, expect } from 'vitest';
import { ringLayout, neighbour, connectors } from './mapLayout.js';

const key = (c) => `${c.col},${c.row}`;

describe('ringLayout', () => {
  it.each([0, 1, 5, 20])('places %i tiles with no overlaps and the centre free', (count) => {
    const layout = ringLayout(count, { reserveBottom: true });
    const taken = [layout.centre, layout.bottom, ...layout.cells].map(key);
    expect(layout.cells).toHaveLength(count);
    expect(new Set(taken).size).toBe(taken.length);
    expect(layout.centre).toEqual({ col: 2, row: 2 });
  });

  it('fills the top row first, deterministically', () => {
    expect(ringLayout(3).cells).toEqual([{ col: 1, row: 1 }, { col: 2, row: 1 }, { col: 3, row: 1 }]);
  });

  it('keeps the bottom middle for the graveyard, below everything else', () => {
    expect(ringLayout(4, { reserveBottom: true }).bottom).toEqual({ col: 2, row: 3 });
    const big = ringLayout(20, { reserveBottom: true });
    expect(big.bottom.row).toBe(Math.max(...big.cells.map((c) => c.row)) + 1);
  });

  it('uses the bottom middle for a tile when nothing is reserved', () => {
    expect(ringLayout(8).cells.map(key)).toContain('2,3');
    expect(ringLayout(8).bottom).toBeNull();
  });
});

describe('neighbour', () => {
  const cells = [{ col: 1, row: 1 }, { col: 2, row: 1 }, { col: 3, row: 1 }, { col: 2, row: 2 }];
  it('moves right, down and left by position', () => {
    expect(neighbour(cells, 0, 'ArrowRight')).toBe(1);
    expect(neighbour(cells, 1, 'ArrowDown')).toBe(3);
    expect(neighbour(cells, 3, 'ArrowUp')).toBe(1);
  });
  it('stays put at an edge or on another key', () => {
    expect(neighbour(cells, 0, 'ArrowLeft')).toBe(0);
    expect(neighbour(cells, 0, 'a')).toBe(0);
  });
});

describe('connectors', () => {
  it('draws from the centre of the middle cell to the centre of each tile', () => {
    expect(connectors({ col: 2, row: 2 }, [{ col: 1, row: 1 }])).toEqual([{ x1: 150, y1: 150, x2: 50, y2: 50 }]);
  });
});

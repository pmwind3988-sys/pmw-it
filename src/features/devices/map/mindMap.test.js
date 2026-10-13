import { describe, it, expect } from 'vitest';
import {
  phoneCurve,
  flowLayout,
  fitsCircle,
  mindMapLayout, curve, nearestInDirection, donutGradient, ringSizeFor,
  BOARD, PHONE_BOARD, RING_PX, EXTRA_SLOTS, EXTRA_PX,
} from './mindMap.js';

/** Centre-to-centre distance in px on a board. */
const px = (a, b, board) => Math.hypot(((a.x - b.x) * board.width) / 100, ((a.y - b.y) * board.height) / 100);

describe('mindMapLayout', () => {
  it.each([1, 2, 3, 4, 5, 6, 7, 8])('keeps %i rings inside the board and clear of each other', (count) => {
    const layout = mindMapLayout({ rings: count, extras: ['stash', 'noLocation', 'retired'] });
    expect(layout.rings).toHaveLength(count);
    for (const { x, y } of [...layout.rings, ...Object.values(layout.extras)]) {
      expect(x).toBeGreaterThanOrEqual(5);
      expect(x).toBeLessThanOrEqual(95);
      expect(y).toBeGreaterThanOrEqual(5);
      expect(y).toBeLessThanOrEqual(95);
    }
    const ringGap = RING_PX[layout.ringSize] * 0.92;
    for (let i = 0; i < count; i += 1) {
      for (let j = i + 1; j < count; j += 1) {
        expect(px(layout.rings[i], layout.rings[j], BOARD)).toBeGreaterThanOrEqual(ringGap);
      }
    }
  });

  it.each([1, 2, 3, 4, 5, 6, 7, 8])('keeps %i rings and the extras apart', (count) => {
    const layout = mindMapLayout({ rings: count, extras: ['stash', 'noLocation', 'retired'] });
    const footprint = { stash: EXTRA_PX.stash, noLocation: EXTRA_PX.noLocation, retired: EXTRA_PX.retired };
    for (const [key, at] of Object.entries(layout.extras)) {
      for (const ring of layout.rings) {
        expect(px(at, ring, BOARD)).toBeGreaterThanOrEqual((footprint[key] / 2 + RING_PX[layout.ringSize] / 2) * 0.9);
      }
    }
  });

  it.each([1, 2, 3, 4])('fits %i rings on a phone board without overlap', (count) => {
    const layout = mindMapLayout({ rings: count, extras: ['stash', 'retired'], ringSize: 'lg' });
    const phoneRing = 104;
    for (let i = 0; i < count; i += 1) {
      for (let j = i + 1; j < count; j += 1) {
        expect(px(layout.rings[i], layout.rings[j], PHONE_BOARD)).toBeGreaterThanOrEqual(phoneRing * 0.92);
      }
    }
  });

  it('is deterministic', () => {
    expect(mindMapLayout({ rings: 5, extras: ['stash'] })).toEqual(mindMapLayout({ rings: 5, extras: ['stash'] }));
  });

  it('only places the extras it is asked for, and ignores unknown names', () => {
    const layout = mindMapLayout({ rings: 2, extras: ['stash', 'bogus'] });
    expect(Object.keys(layout.extras)).toEqual(['stash']);
    expect(layout.extras.stash).toEqual(EXTRA_SLOTS.stash);
  });

  it('shrinks the rings as the count grows', () => {
    expect(ringSizeFor(2)).toBe('lg');
    expect(ringSizeFor(4)).toBe('md');
    expect(ringSizeFor(8)).toBe('sm');
    expect(mindMapLayout({ rings: 8 }).ringSize).toBe('sm');
  });

  it('accepts a forced ring size', () => {
    expect(mindMapLayout({ rings: 2, ringSize: 'md' }).ringSize).toBe('md');
  });
});

describe('curve', () => {
  it('returns a quadratic path ending at the target point', () => {
    expect(curve({ x: 50, y: 48 }, { x: 50, y: 48 })).toBe('M 50 48 Q 50 48 50 48');
    expect(curve({ x: 50, y: 48 }, { x: 20, y: 30 })).toMatch(/^M 50 48 Q [-\d.]+ [-\d.]+ 20 30$/);
  });
});

describe('nearestInDirection', () => {
  const points = [
    { x: 20, y: 30 }, // 0: upper left
    { x: 80, y: 30 }, // 1: upper right
    { x: 50, y: 85 }, // 2: bottom
  ];

  it('moves to the nearest node that way', () => {
    expect(nearestInDirection(points, 0, 'ArrowRight')).toBe(1);
    expect(nearestInDirection(points, 1, 'ArrowLeft')).toBe(0);
    expect(nearestInDirection(points, 0, 'ArrowDown')).toBe(2);
  });

  it('stays put when nothing lies that way, or the key is not an arrow', () => {
    expect(nearestInDirection(points, 0, 'ArrowUp')).toBe(0);
    expect(nearestInDirection(points, 0, 'ArrowLeft')).toBe(0);
    expect(nearestInDirection(points, 0, 'Enter')).toBe(0);
  });
});

describe('donutGradient', () => {
  it('draws one slice per grade present, in severity order, from the grade variables', () => {
    const css = donutGradient({ Critical: 1, 'Needs attention': 1, Moderate: 0, Optimal: 2, Unknown: 0 }, 4);
    expect(css).toBe('conic-gradient(var(--grade-critical) 0% 25%, var(--grade-attention) 25% 50%, var(--grade-optimal) 50% 100%)');
  });

  it('is a plain grey ring for a place with no machines', () => {
    expect(donutGradient({}, 0)).toBe('var(--grade-unknown)');
  });
});

describe('flowLayout', () => {
  const nodes = (rings, extras = []) => [
    ...Array.from({ length: rings }, (_, i) => ({ key: `r${i}`, kind: 'dept' })),
    ...extras.map((kind) => ({ key: kind, kind })),
  ];
  const px = (a, b, width) => Math.hypot(((a.x - b.x) * width) / 100, a.y - b.y);

  it.each([
    [12, 900, 140, false], [12, 700, 140, false], [8, 343, 104, true], [3, 343, 104, true], [20, 1200, 140, false],
  ])('keeps %i rings and their captions apart on a %ipx board', (count, width, ringPx, phone) => {
    const layout = flowLayout(nodes(count, ['stash', 'retired']), { width, ringPx, phone });
    const points = Object.values(layout.at);
    for (let i = 0; i < points.length; i += 1) {
      for (let j = i + 1; j < points.length; j += 1) {
        const sameRow = points[i].y === points[j].y;
        // Side by side: a ring each plus a gap. One above the other: a ring plus its caption.
        const need = sameRow ? ringPx + 8 : ringPx + 90;
        expect(px(points[i], points[j], width)).toBeGreaterThanOrEqual(need);
      }
    }
    for (const point of points) {
      expect(point.x).toBeGreaterThan(0);
      expect(point.x).toBeLessThan(100);
      expect(point.y).toBeLessThan(layout.height);
    }
  });

  it('puts the hub above every ring and grows the board with the count', () => {
    const few = flowLayout(nodes(3), { width: 700 });
    const many = flowLayout(nodes(15), { width: 700 });
    expect(many.height).toBeGreaterThan(few.height);
    expect(Math.min(...Object.values(few.at).map((p) => p.y)) - few.hub.y).toBeGreaterThan(120);
  });

  it('hangs the first row off the hub and every later node off the row above', () => {
    const layout = flowLayout(nodes(7, ['stash']), { width: 700 });
    const firstRowY = layout.at.r0.y;
    for (const [key, point] of Object.entries(layout.at)) {
      const up = layout.parent[key];
      if (point.y === firstRowY) expect(up).toBeNull();
      else expect(layout.at[up].y).toBeLessThan(point.y);
    }
  });

  it('puts the extras in their own row after the rings', () => {
    const layout = flowLayout(nodes(5, ['stash', 'noLocation', 'retired']), { width: 900 });
    expect(layout.at.stash.y).toBe(layout.at.retired.y);
    expect(layout.at.stash.y).toBeGreaterThan(layout.at.r4.y);
  });
});

describe('fitsCircle', () => {
  it('keeps the circle only for a few rings on a wide board', () => {
    expect(fitsCircle(3, 1000)).toBe(true);
    expect(fitsCircle(5, 1000)).toBe(false);
    expect(fitsCircle(3, 600)).toBe(false);
  });
});

describe('phoneCurve', () => {
  it('stays between its two ends across the board', () => {
    const d = phoneCurve({ x: 50, y: 64 }, { x: 18, y: 760 });
    const xs = d.match(/-?\d+(\.\d+)?/g).map(Number).filter((_, i) => i % 2 === 0);
    expect(Math.min(...xs)).toBeGreaterThanOrEqual(18);
    expect(Math.max(...xs)).toBeLessThanOrEqual(50);
  });
});

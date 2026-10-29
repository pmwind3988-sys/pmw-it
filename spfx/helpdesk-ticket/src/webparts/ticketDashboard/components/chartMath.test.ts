import {
  anchorShift, axisTicks, barPct, heatColour, labelEvery, moveIndex, niceMax, percentText, PALETTE,
  snapIndex, svgArea, svgPoints, svgY, ticketsText, xPct, yPct
} from './chartMath';

describe('niceMax', () => {
  it('rounds up to the nice steps', () => {
    expect(niceMax(0)).toBe(5);
    expect(niceMax(-3)).toBe(5);
    expect(niceMax(5)).toBe(5);
    expect(niceMax(6)).toBe(10);
    expect(niceMax(21)).toBe(25);
    expect(niceMax(54)).toBe(100);
    expect(niceMax(91)).toBe(100);
  });

  it('carries on through the decades', () => {
    expect(niceMax(101)).toBe(200);
    expect(niceMax(251)).toBe(500);
    expect(niceMax(1200)).toBe(2000);
  });

  it('treats NaN as nothing to plot', () => {
    expect(niceMax(NaN)).toBe(5);
  });
});

describe('axisTicks', () => {
  it('gives four or five whole-number steps from zero to the top', () => {
    expect(axisTicks(5)).toEqual([0, 1, 2, 3, 4, 5]);
    expect(axisTicks(10)).toEqual([0, 2, 4, 6, 8, 10]);
    expect(axisTicks(20)).toEqual([0, 5, 10, 15, 20]);
    expect(axisTicks(100)).toEqual([0, 25, 50, 75, 100]);
  });

  it('only ever yields whole numbers for the nice maxima', () => {
    [5, 10, 20, 25, 50, 100, 200, 250, 500].forEach((max) => {
      axisTicks(max).forEach((t) => expect(Number.isInteger(t)).toBe(true));
    });
  });
});

describe('labelEvery', () => {
  it('shows every label when there are few', () => {
    expect(labelEvery(0)).toBe(1);
    expect(labelEvery(12)).toBe(1);
  });

  it('thins the labels so no more than twelve show', () => {
    expect(labelEvery(13)).toBe(2);
    expect(labelEvery(31)).toBe(3);
    expect(labelEvery(400)).toBe(34);
    expect(Math.ceil(400 / labelEvery(400))).toBeLessThanOrEqual(12);
  });
});

describe('percentText and barPct', () => {
  it('reads a share in whole percent', () => {
    expect(percentText(1, 3)).toBe('33%');
    expect(percentText(2, 3)).toBe('67%');
    expect(percentText(1, 1)).toBe('100%');
  });

  it('reads zero when there is no whole to share', () => {
    expect(percentText(0, 0)).toBe('0%');
    expect(percentText(4, 0)).toBe('0%');
  });

  it('scales a bar against the largest value and never past full', () => {
    expect(barPct(5, 10)).toBe(50);
    expect(barPct(12, 10)).toBe(100);
    expect(barPct(0, 10)).toBe(0);
    expect(barPct(3, 0)).toBe(0);
    expect(barPct(NaN, 10)).toBe(0);
  });
});

describe('positions', () => {
  it('puts a lone point in the middle and spreads the rest edge to edge', () => {
    expect(xPct(0, 1)).toBe(50);
    expect(xPct(0, 3)).toBe(0);
    expect(xPct(1, 3)).toBe(50);
    expect(xPct(2, 3)).toBe(100);
  });

  it('reads zero at the baseline and the top at the top', () => {
    expect(yPct(0, 100)).toBe(100);
    expect(yPct(100, 100)).toBe(0);
    expect(svgY(50, 100)).toBe(150);
  });

  it('builds the polyline and area in viewBox units', () => {
    expect(svgPoints([0, 50], 100)).toBe('0,300 1000,150');
    expect(svgArea([0, 50], 100)).toBe('M0,300 L0,300 1000,150 L1000,300 Z');
  });

  it('keeps end labels inside the plot', () => {
    expect(anchorShift(0)).toBe('0%');
    expect(anchorShift(100)).toBe('-100%');
    expect(anchorShift(50)).toBe('-50%');
  });
});

describe('hover and keyboard', () => {
  it('snaps the pointer to the nearest point', () => {
    expect(snapIndex(0, 5)).toBe(0);
    expect(snapIndex(0.5, 5)).toBe(2);
    expect(snapIndex(1, 5)).toBe(4);
    expect(snapIndex(0.3, 5)).toBe(1);
    expect(snapIndex(2, 5)).toBe(4);
    expect(snapIndex(NaN, 5)).toBe(0);
    expect(snapIndex(0.7, 1)).toBe(0);
  });

  it('moves one point at a time and stops at either end', () => {
    expect(moveIndex(undefined, 5, -1)).toBe(3);
    expect(moveIndex(undefined, 5, 1)).toBe(4);
    expect(moveIndex(0, 5, -1)).toBe(0);
    expect(moveIndex(4, 5, 1)).toBe(4);
    expect(moveIndex(2, 5, 1)).toBe(3);
    expect(moveIndex(2, 0, 1)).toBe(0);
  });
});

describe('heatColour', () => {
  const ramp = PALETTE.heatRamp;

  it('is the empty colour for a zero cell', () => {
    expect(heatColour(0, 10, ramp, 'empty')).toBe('empty');
    expect(heatColour(3, 0, ramp, 'empty')).toBe('empty');
  });

  it('steps up the ramp with the count and tops out at the darkest', () => {
    expect(heatColour(1, 10, ramp, 'empty')).toBe(ramp[0]);
    expect(heatColour(6, 10, ramp, 'empty')).toBe(ramp[3]);
    expect(heatColour(10, 10, ramp, 'empty')).toBe(ramp[4]);
  });
});

describe('ticketsText', () => {
  it('pluralises', () => {
    expect(ticketsText(1)).toBe('1 ticket');
    expect(ticketsText(0)).toBe('0 tickets');
    expect(ticketsText(12)).toBe('12 tickets');
  });
});

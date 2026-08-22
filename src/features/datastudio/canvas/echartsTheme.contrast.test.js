import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

// Spec §10.7 requires the series palette to be colour-blind-safe and
// legible on BOTH panel colours. The plan says to verify that rather
// than assume it, so this reads the real tokens out of the stylesheet
// and checks them numerically. Eyeballing a palette is exactly how one
// ships two greens that a deuteranope cannot tell apart.
//
// The stylesheet is parsed as text rather than imported, because these
// are CSS custom properties and there is no DOM here to compute them
// against.

const CSS = readFileSync(
  fileURLToPath(new URL('../../../styles/datastudio.css', import.meta.url)),
  'utf8',
);

function paletteFrom(blockStart) {
  const start = CSS.indexOf(blockStart);
  const block = CSS.slice(start, CSS.indexOf('}', start));
  return Array.from({ length: 8 }, (_, i) => {
    const match = new RegExp(`--ds-series-${i + 1}:\\s*(#[0-9a-fA-F]{6})`).exec(block);
    return match?.[1];
  });
}

const LIGHT = paletteFrom(':root {\n  --ds-series-1');
const DARK = paletteFrom("[data-theme='dark'] {\n  --ds-series-1");

// The --it-panel each palette is drawn on, from shell.css.
const LIGHT_PANEL = '#ffffff';
const DARK_PANEL = '#151b24';

function rgb(hex) {
  return [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
}

function relativeLuminance(hex) {
  const [r, g, b] = rgb(hex).map((c) => (
    c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

function contrast(a, b) {
  const la = relativeLuminance(a);
  const lb = relativeLuminance(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

// Brettel-style simulation, simplified: the standard LMS-projection
// matrices for full dichromacy. Good enough to answer the only question
// asked here -- do two palette entries collapse onto each other.
const CVD_MATRICES = {
  protanopia: [[0.567, 0.433, 0], [0.558, 0.442, 0], [0, 0.242, 0.758]],
  deuteranopia: [[0.625, 0.375, 0], [0.7, 0.3, 0], [0, 0.3, 0.7]],
  tritanopia: [[0.95, 0.05, 0], [0, 0.433, 0.567], [0, 0.475, 0.525]],
};

function simulate(hex, kind) {
  const [r, g, b] = rgb(hex);
  const m = CVD_MATRICES[kind];
  return [
    m[0][0] * r + m[0][1] * g + m[0][2] * b,
    m[1][0] * r + m[1][1] * g + m[1][2] * b,
    m[2][0] * r + m[2][1] * g + m[2][2] * b,
  ];
}

function distance(a, b) {
  return Math.sqrt(a.reduce((sum, v, i) => sum + (v - b[i]) ** 2, 0));
}

describe('series palette', () => {
  it('defines eight colours in both themes', () => {
    expect(LIGHT.filter(Boolean)).toHaveLength(8);
    expect(DARK.filter(Boolean)).toHaveLength(8);
  });

  it('uses a different palette per theme', () => {
    expect(LIGHT).not.toEqual(DARK);
  });

  // 3:1 is the WCAG threshold for non-text graphical objects, which is
  // what a bar or a line is.
  it.each([0, 1, 2, 3, 4, 5, 6, 7])(
    'light slot %i clears 3:1 against the light panel', (i) => {
      expect(contrast(LIGHT[i], LIGHT_PANEL)).toBeGreaterThanOrEqual(3);
    });

  it.each([0, 1, 2, 3, 4, 5, 6, 7])(
    'dark slot %i clears 3:1 against the dark panel', (i) => {
      expect(contrast(DARK[i], DARK_PANEL)).toBeGreaterThanOrEqual(3);
    });

  // The first four slots carry most charts. They are the ones that must
  // stay apart under colour-blindness; slots 5-8 only appear on charts
  // busy enough that a legend is doing the work anyway.
  describe.each(['protanopia', 'deuteranopia', 'tritanopia'])('under %s', (kind) => {
    it.each([['light', () => LIGHT], ['dark', () => DARK]])(
      'keeps the first four %s slots distinguishable', (_name, get) => {
        const palette = get().slice(0, 4);
        for (let a = 0; a < palette.length; a++) {
          for (let b = a + 1; b < palette.length; b++) {
            const d = distance(simulate(palette[a], kind), simulate(palette[b], kind));
            expect(d).toBeGreaterThan(0.12);
          }
        }
      });
  });
});

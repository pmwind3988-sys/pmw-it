import { GRADES, UNKNOWN } from './defaultStandard.js';

/** The CSS name of each grade: `.g-critical`, `--grade-critical`, … */
export const GRADE_SLUG = {
  Critical: 'critical', 'Needs attention': 'attention', Moderate: 'moderate', Optimal: 'optimal', Unknown: 'unknown',
};

const DARK = '#101828';
const LIGHT = '#ffffff';

function rgb(hex) {
  const n = parseInt(String(hex).slice(1), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function luminance(hex) {
  const [r, g, b] = rgb(hex);
  return [r / 255, g / 255, b / 255]
    .map((v, i) => v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4)
    .reduce((sum, c, i) => sum + c * [0.2126, 0.7152, 0.0722][i], 0);
}

export function contrastRatio(a, b) {
  const [x, y] = [luminance(a), luminance(b)];
  return (Math.max(x, y) + 0.05) / (Math.min(x, y) + 0.05);
}

/** Text on a grade chip: whichever of white or near-black reads better. */
export function inkFor(hex) {
  return contrastRatio(hex, LIGHT) >= contrastRatio(hex, DARK) ? LIGHT : DARK;
}

/** Perceptual colour distance using redmean algorithm. */
function colourDistance(a, b) {
  const [r1, g1, b1] = rgb(a);
  const [r2, g2, b2] = rgb(b);
  const rm = (r1 + r2) / 2;
  return Math.sqrt((2 + rm / 256) * (r1 - r2) ** 2 + 4 * (g1 - g2) ** 2 + (2 + (255 - rm) / 256) * (b1 - b2) ** 2);
}

/** Two of the four grades too alike to tell apart (colour distance under 100). */
export function tooSimilar(colors) {
  for (let i = 0; i < GRADES.length; i += 1) {
    for (let j = i + 1; j < GRADES.length; j += 1) {
      if (colourDistance(colors[GRADES[i]], colors[GRADES[j]]) < 100) return true;
    }
  }
  return false;
}

export function gradeCssVars(colors) {
  const vars = {};
  for (const grade of [...GRADES, UNKNOWN]) {
    const slug = GRADE_SLUG[grade];
    vars[`--grade-${slug}`] = colors[grade];
    vars[`--grade-${slug}-ink`] = inkFor(colors[grade]);
  }
  return vars;
}

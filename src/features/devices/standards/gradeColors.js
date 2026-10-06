import { GRADES, UNKNOWN } from './defaultStandard.js';

/** The CSS name of each grade: `.g-critical`, `--grade-critical`, … */
export const GRADE_SLUG = {
  Critical: 'critical', 'Needs attention': 'attention', Moderate: 'moderate', Optimal: 'optimal', Unknown: 'unknown',
};

const DARK = '#101828';
const LIGHT = '#ffffff';

function luminance(hex) {
  const n = parseInt(String(hex).slice(1), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255]
    .map((v) => { const c = v / 255; return c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4; })
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

/** Two of the four grades too alike to tell apart (contrast under 1.5). */
export function tooSimilar(colors) {
  for (let i = 0; i < GRADES.length; i += 1) {
    for (let j = i + 1; j < GRADES.length; j += 1) {
      if (contrastRatio(colors[GRADES[i]], colors[GRADES[j]]) < 1.5) return true;
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

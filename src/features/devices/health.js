import { GRADES, UNKNOWN, PARTS } from './standards/defaultStandard.js';

/**
 * The redesign's one health language is the part grade. These two helpers are
 * what every screen reads it through, so a ring on the map, a row in the list
 * and a card in a department cannot disagree about a machine.
 */

/** The worst grade any of a machine's five parts holds; Unknown when none was graded. */
export function worstGradeOf(device) {
  const held = new Set(PARTS.map(({ key }) => device?.parts?.[key]?.grade));
  return GRADES.find((grade) => held.has(grade)) ?? UNKNOWN;
}

/** Critical or Needs attention on at least one part: what "needs work" means everywhere. */
export const needsWork = (device) => {
  const worst = worstGradeOf(device);
  return worst === 'Critical' || worst === 'Needs attention';
};

/** Severity order for sorting, worst first; Unknown sorts after Optimal. */
export const gradeRank = (grade) => {
  const at = [...GRADES, UNKNOWN].indexOf(grade);
  return at === -1 ? GRADES.length + 1 : at;
};

/**
 * The risk score keeps its job but loses the word "Critical", which the part
 * grades already use for something else. Stored values are unchanged
 * (Critical / High / Watch / OK); this is only what a person reads.
 */
const SECURITY_RISK = { Critical: 'High', High: 'High', Watch: 'Medium', OK: 'Low' };

export const securityRiskOf = (riskLevel) => SECURITY_RISK[riskLevel] ?? 'Unknown';

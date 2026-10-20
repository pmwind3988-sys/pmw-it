/**
 * Which colour family a category's logo is drawn in.
 *
 * Five families rather than fifteen colours: a register where every row is a
 * different hue reads as confetti, and the family is what the eye uses to find
 * "the cables" in a long list. A category somebody added sits with
 * "everything else" until it is given a place here.
 */

const FAMILIES = {
  Laptop: 'comp',
  Desktop: 'comp',
  Monitor: 'comp',
  Phone: 'comp',
  Tab: 'comp',
  Printer: 'office',
  'Docking Station': 'office',
  Network: 'office',
  Keyboard: 'input',
  Mouse: 'input',
  'PC Part': 'parts',
  Cable: 'parts',
  Adapter: 'parts',
};

export function familyOf(category) {
  return FAMILIES[category] ?? 'misc';
}

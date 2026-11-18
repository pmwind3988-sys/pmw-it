/**
 * Which picture a category gets. Built-in categories by name; one somebody
 * added at the till ("Headsets", "USB-C hub") by the words in it; else a box.
 * Pure, and kept apart from the component so Fast Refresh keeps working.
 */
const RULES = [
  ['dock', /dock|hub/],
  // Before desktop: "PC Part" holds "pc".
  ['part', /pc part|\bram\b|ssd|hdd|memory|\bpart/],
  ['laptop', /laptop|notebook/],
  ['desktop', /desktop|\bpc\b|tower|workstation/],
  ['part', /pc part|\bram\b|ssd|hdd|memory|\bpart/],
  ['monitor', /monitor|screen|display/],
  ['printer', /printer|scanner/],
  ['tablet', /\btab\b|tablet|ipad/],
  ['phone', /phone|mobile|sim/],
  ['network', /network|router|switch|wifi|access point/],
  ['keyboard', /keyboard/],
  ['mouse', /mouse|mice/],
  ['cable', /cable|hdmi|vga|cord|lan\b/],
  ['adapter', /adapter|adaptor|charger|power/],
  ['headset', /headset|headphone|earphone|speaker|webcam/],
  ['accessory', /accessor|bag|stand/],
];

export function artKeyFor(category) {
  const name = String(category ?? '').toLowerCase();
  return RULES.find(([, pattern]) => pattern.test(name))?.[0] ?? 'other';
}

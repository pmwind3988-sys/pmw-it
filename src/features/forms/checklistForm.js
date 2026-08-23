/**
 * The asset checklist, declared as data.
 *
 * This is the employee's own signed record of what they received or handed
 * back — not a request, and not connected to the asset register, which is IT's
 * own. It follows the IT ASSET TRACKING FORM supplied as the reference: its
 * fields, its order, and its branching.
 *
 * Kept apart from the page so that "an OUT checklist needs a signature" and
 * "an individual request needs at least one item" can be tested without
 * rendering anything. Those rules are the part that can be wrong without
 * looking wrong.
 */

/**
 * What the form READS and what it STORES are deliberately different.
 *
 * The reference form says IN / OUT / INDIVIDUAL REQUEST. The list already holds
 * `In` / `Out` / `Individual Request` on every checklist ever signed, and
 * changing the stored strings would orphan all of them. So the label is the
 * reference's and the value is the list's, mapped in one place.
 */
export const FORM_MODES = [
  {
    value: 'In',
    label: 'IN',
    description: 'New Joiner Provisioning',
  },
  {
    value: 'Out',
    label: 'OUT',
    description: 'Exit / Offboarding',
  },
  {
    value: 'Individual Request',
    label: 'INDIVIDUAL REQUEST',
    description: '(Ad-Hoc) Request',
  },
];

export const IN = 'In';
export const OUT = 'Out';
export const INDIVIDUAL = 'Individual Request';

export const ENTITIES = ['PMW', 'PCI', 'PML', 'WB'];

/**
 * What an IN or OUT checklist ticks off. Shorter than the request list below
 * and that is the reference form's own distinction, not an oversight: a
 * handover checklist covers what a person is issued, where an ad-hoc request
 * can also be for a cable or a spare desktop.
 */
export const CHECKLIST_ITEMS = [
  'Laptop',
  'Mouse',
  'Monitor',
  'Keyboard',
  'Speaker',
  'Earphone',
  'Phone & Simcard',
  'Locker Key',
];

export const REQUESTABLE_ITEMS = [
  'Laptop',
  'Desktop',
  'Mouse',
  'Monitor',
  'Keyboard',
  'HDMI Cable',
  'VGA Cable',
  'Speaker',
  'Earphone',
  'Phone & Simcard',
  'Locker Key',
];

export const modeLabel = (value) =>
  FORM_MODES.find((mode) => mode.value === value)?.label ?? value ?? '';

/** IN and OUT tick a list; an individual request names items and quantities. */
export const isRequest = (mode) => mode === INDIVIDUAL;

export function newItemRow() {
  return { item: '', quantity: 1 };
}

export function emptyChecklist() {
  return {
    formMode: '',
    employeeName: '',
    employeeNo: '',
    position: '',
    entity: '',
    // The date the employee says it happened on, which is theirs to change.
    // The instant it was actually submitted is stamped separately at save.
    formDate: todayValue(),
    checkedItems: [],
    items: [newItemRow()],
    serialNumbers: '',
    otherRemarks: '',
    signature: null,
  };
}

/** `yyyy-mm-dd` in the LOCAL day, not UTC — near midnight those differ. */
export function todayValue(now = new Date()) {
  const offset = now.getTimezoneOffset() * 60000;
  return new Date(now.getTime() - offset).toISOString().slice(0, 10);
}

/**
 * The fields step two shows, given the mode. The page renders whatever this
 * returns rather than deciding for itself, so the branching is one testable
 * answer instead of conditions scattered through JSX.
 */
export const CHECKLIST_STEPS = ['Form type', 'Details'];

export function fieldsFor(mode) {
  const common = ['employeeName', 'employeeNo', 'position', 'entity', 'formDate'];
  const closing = ['serialNumbers', 'otherRemarks', 'signature'];

  if (isRequest(mode)) return [...common, 'items', ...closing];
  return [...common, 'checkedItems', ...closing];
}

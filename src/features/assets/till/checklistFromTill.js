import {
  INDIVIDUAL, OUT, CHECKLIST_ITEMS, REQUESTABLE_ITEMS, emptyChecklist, todayValue,
} from '../../forms/checklistForm.js';

/**
 * A till receipt as an asset checklist: the employee's own signed record of
 * what they received (IN, INDIVIDUAL REQUEST) or handed back (OUT).
 *
 * The checklist has a fixed tick-list (IN/OUT) or "item × quantity" lines
 * (an individual request), and ONE box for serial numbers. So each scanned
 * thing is ticked or listed where the checklist has a place for it; every
 * serial goes in the serial box, one line per item; and anything the
 * checklist has no place for goes in "Other remarks" with its name and serial
 * — nothing scanned is left off the form the employee signs.
 *
 * Employee no., entity and department are not known to the till and are left
 * blank, which is what lets the employee fill them in.
 */

/** Where a till category goes on the checklist, if anywhere. */
export function checklistItemFor(row, formMode) {
  const category = String(row?.category ?? '');
  const name = String(row?.name ?? '');
  const request = formMode === INDIVIDUAL;
  const offered = request ? REQUESTABLE_ITEMS : CHECKLIST_ITEMS;

  let item = null;
  if (['Laptop', 'Mouse', 'Monitor', 'Keyboard'].includes(category)) item = category;
  else if (category === 'Desktop') item = 'Desktop';
  else if (category === 'Phone') item = 'Phone & Simcard';
  else if (category === 'Cable' && /\bhdmi\b/i.test(name)) item = 'HDMI Cable';
  else if (category === 'Cable' && /\bvga\b/i.test(name)) item = 'VGA Cable';

  return item && offered.includes(item) ? item : null;
}

const plural = (n) => (n > 1 ? ` × ${n}` : '');

/**
 * `rows`: [{ category, name, serial, qty }] — what the till handed out or took
 * back. `person`: { name, title }. Returns checklist values, ready for
 * `draftLink`.
 */
export function checklistFromTill({ formMode, person = {}, rows = [], date = todayValue() }) {
  const values = {
    ...emptyChecklist(),
    formMode,
    employeeName: String(person.name ?? '').trim(),
    position: String(person.title ?? '').trim(),
    formDate: date,
  };

  const ticked = new Set();
  const requested = new Map();
  const remarks = [];
  const serials = [];

  for (const row of rows) {
    const qty = Math.max(1, Number(row.qty) || 1);
    const name = String(row.name ?? '').trim() || String(row.category ?? 'Item');
    const serial = String(row.serial ?? '').trim();
    const item = checklistItemFor(row, formMode);

    if (serial) serials.push(`${row.category || 'Item'} · ${name} · S/N ${serial}`);

    if (!item) {
      remarks.push(`${name}${plural(qty)}${serial ? ` · S/N ${serial}` : ''}`);
    } else if (formMode === INDIVIDUAL) {
      requested.set(item, (requested.get(item) ?? 0) + qty);
    } else {
      ticked.add(item);
    }
  }

  if (formMode === INDIVIDUAL) {
    values.items = [...requested].map(([item, quantity]) => ({ item, quantity }));
    if (!values.items.length) values.items = emptyChecklist().items;
  } else {
    values.checkedItems = CHECKLIST_ITEMS.filter((item) => ticked.has(item));
  }

  values.serialNumbers = serials.join('\n');
  if (remarks.length) {
    values.otherRemarks = `${formMode === OUT ? 'Also returned' : 'Also handed over'}:\n${remarks.join('\n')}`;
  }
  return values;
}

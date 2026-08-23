import { describe, it, expect } from 'vitest';
import {
  FORM_MODES, ENTITIES, CHECKLIST_ITEMS, REQUESTABLE_ITEMS,
  fieldsFor, isRequest, modeLabel, emptyChecklist, todayValue,
  IN, OUT, INDIVIDUAL,
} from './checklistForm.js';
import { validateChecklist, validateRequest, hasErrors } from './validate.js';
import {
  toChecklistItem, formatItems, formatChecked, parseFormDate, signatureFileName,
} from './toChecklistItem.js';
import {
  newEmployee, employeeFromItem, employeeToItem, dateLabel, parseDay, MAX_EMPLOYEES,
} from './requestForm.js';

const complete = (overrides = {}) => ({
  ...emptyChecklist(),
  formMode: IN,
  employeeName: 'Amir Hakim',
  employeeNo: 'E-1042',
  position: 'Engineer',
  entity: 'PMW',
  signature: 'data:image/png;base64,AAA',
  ...overrides,
});

describe('the declaration follows the reference form', () => {
  it('offers the three form types, read as the reference reads them', () => {
    expect(FORM_MODES.map((mode) => mode.label))
      .toEqual(['IN', 'OUT', 'INDIVIDUAL REQUEST']);
  });

  /**
   * The list already holds these on every checklist ever signed. Changing the
   * stored strings would orphan all of them.
   */
  it('stores the values the list already has', () => {
    expect(FORM_MODES.map((mode) => mode.value))
      .toEqual(['In', 'Out', 'Individual Request']);
  });

  it('offers the reference form entities', () => {
    expect(ENTITIES).toEqual(['PMW', 'PCI', 'PML', 'WB']);
  });

  it('ticks a shorter list than it can request from', () => {
    expect(CHECKLIST_ITEMS).toEqual([
      'Laptop', 'Mouse', 'Monitor', 'Keyboard', 'Speaker', 'Earphone',
      'Phone & Simcard', 'Locker Key',
    ]);
    expect(REQUESTABLE_ITEMS).toContain('Desktop');
    expect(REQUESTABLE_ITEMS).toContain('HDMI Cable');
    expect(CHECKLIST_ITEMS).not.toContain('Desktop');
  });

  it('reads a mode back as its label', () => {
    expect(modeLabel(INDIVIDUAL)).toBe('INDIVIDUAL REQUEST');
    expect(modeLabel('nonsense')).toBe('nonsense');
  });
});

describe('branching', () => {
  it('shows IN and OUT the tick list', () => {
    expect(fieldsFor(IN)).toContain('checkedItems');
    expect(fieldsFor(OUT)).toContain('checkedItems');
    expect(fieldsFor(IN)).not.toContain('items');
  });

  it('shows an individual request the item rows instead', () => {
    expect(fieldsFor(INDIVIDUAL)).toContain('items');
    expect(fieldsFor(INDIVIDUAL)).not.toContain('checkedItems');
  });

  it('asks all three for the same details and the same closing fields', () => {
    for (const mode of [IN, OUT, INDIVIDUAL]) {
      expect(fieldsFor(mode)).toEqual(expect.arrayContaining([
        'employeeName', 'employeeNo', 'position', 'entity', 'formDate',
        'serialNumbers', 'otherRemarks', 'signature',
      ]));
    }
  });

  it('knows which mode is a request', () => {
    expect(isRequest(INDIVIDUAL)).toBe(true);
    expect(isRequest(IN)).toBe(false);
  });
});

describe('validateChecklist', () => {
  it('is happy with a complete IN checklist', () => {
    expect(validateChecklist(complete())).toEqual({});
  });

  /** Marking the whole form red before step 2 was even reached is hostile. */
  it('asks only for the form type on step one', () => {
    expect(validateChecklist(emptyChecklist(), { step: 0 }))
      .toEqual({ formMode: 'Form type is required.' });
  });

  it('lets step one pass once a type is picked', () => {
    expect(validateChecklist({ ...emptyChecklist(), formMode: IN }, { step: 0 })).toEqual({});
  });

  /** All at once — being told about one field at a time is the worst version. */
  it('reports every missing field together', () => {
    const errors = validateChecklist({ ...emptyChecklist(), formMode: IN });

    expect(Object.keys(errors)).toEqual(expect.arrayContaining([
      'employeeName', 'employeeNo', 'position', 'entity', 'signature',
    ]));
  });

  it('treats whitespace as empty', () => {
    expect(validateChecklist(complete({ employeeName: '   ' })).employeeName).toBeDefined();
  });

  /** The signature is what makes this a record rather than a note. */
  it('refuses an unsigned checklist', () => {
    expect(validateChecklist(complete({ signature: null })).signature).toBeDefined();
  });

  describe('an individual request', () => {
    const request = (items) => complete({ formMode: INDIVIDUAL, items });

    it('needs at least one item', () => {
      expect(validateChecklist(request([{ item: '', quantity: 1 }])).items)
        .toBe('Add at least one item.');
    });

    it('is satisfied by one named item', () => {
      expect(validateChecklist(request([{ item: 'Laptop', quantity: 1 }]))).toEqual({});
    });

    /** A quantity with no item is not a request for anything. */
    it('ignores an empty row beside a filled one', () => {
      expect(validateChecklist(request([
        { item: 'Laptop', quantity: 1 },
        { item: '', quantity: 3 },
      ]))).toEqual({});
    });

    it('refuses a quantity of nothing', () => {
      expect(validateChecklist(request([{ item: 'Laptop', quantity: 0 }])).items)
        .toContain('quantity');
    });

    it('does not ask an IN checklist for items', () => {
      expect(validateChecklist(complete({ items: [] })).items).toBeUndefined();
    });
  });
});

describe('validateRequest', () => {
  const employee = (overrides = {}) => ({
    fullName: 'Amir Hakim',
    position: 'Engineer',
    entity: 'PMW',
    department: 'Engineering',
    joinDate: '2026-09-01',
    ...overrides,
  });

  it('is happy with a complete employee', () => {
    expect(validateRequest([employee()])).toEqual({});
  });

  /** So the right panel is marked, not the whole form. */
  it('keys the problem to the employee who has it', () => {
    const errors = validateRequest([employee(), employee({ fullName: '' })]);

    expect(errors[0]).toBeUndefined();
    expect(errors[1].fullName).toBeDefined();
  });

  it('refuses a submission with nobody in it', () => {
    expect(validateRequest([]).form).toBeDefined();
    expect(hasErrors(validateRequest([]))).toBe(true);
  });

  it('does not require the optional fields', () => {
    expect(validateRequest([employee({ callingName: '', employeeId: '' })])).toEqual({});
  });
});

describe('toChecklistItem', () => {
  /** SharePoint is read by people; JSON in a cell is not a record. */
  it('writes ticked items as readable lines', () => {
    const item = toChecklistItem(complete({ checkedItems: ['Laptop', 'Mouse'] }));

    expect(item.AssetMatrix).toBe('Laptop\nMouse');
    expect(item.RequestedItems).toBe('');
  });

  it('writes requested items with their quantities', () => {
    const item = toChecklistItem(complete({
      formMode: INDIVIDUAL,
      items: [{ item: 'Laptop', quantity: 1 }, { item: 'Monitor', quantity: 2 }],
    }));

    expect(item.RequestedItems).toBe('Laptop x 1\nMonitor x 2');
    expect(item.AssetMatrix).toBe('');
  });

  /** Otherwise re-reading a row shows a previous shape's leftovers. */
  it('clears the shape that does not apply rather than omitting it', () => {
    const item = toChecklistItem(complete({ formMode: INDIVIDUAL, checkedItems: ['Laptop'] }));
    expect(item.AssetMatrix).toBe('');
  });

  it('keeps both dates, because they answer different questions', () => {
    const item = toChecklistItem(complete({ formDate: '2026-08-20' }), { submittedAt: 1755950400000 });

    expect(item.SubmissionDate).toBe(new Date(1755950400000).toISOString());
    expect(item.FormDate).toContain('2026-08-20');
    expect(item.SubmissionDateMYT).toMatch(/(AM|PM)/);
  });

  it('names the row so it is identifiable in the list', () => {
    expect(toChecklistItem(complete()).Title).toBe('Amir Hakim — E-1042');
  });

  it('never leaves the row nameless', () => {
    expect(toChecklistItem(complete({ employeeName: '', employeeNo: '' })).Title)
      .toBe('Asset checklist');
  });

  it('omits an entity nobody picked rather than sending an empty choice', () => {
    expect('Entity' in toChecklistItem(complete({ entity: '' }))).toBe(false);
  });

  it('omits an unparseable date rather than sending "Invalid Date"', () => {
    expect('FormDate' in toChecklistItem(complete({ formDate: 'someday' }))).toBe(false);
  });
});

describe('parseFormDate', () => {
  /**
   * `new Date('2026-08-23')` is UTC midnight, which in Malaysia is the previous
   * day at 8am — so a date picked as the 23rd would store as the 22nd.
   */
  it('reads a picked day as that day, not as the one before', () => {
    const parsed = new Date(parseFormDate('2026-08-23'));

    expect(parsed.getFullYear()).toBe(2026);
    expect(parsed.getMonth()).toBe(7);
    expect(parsed.getDate()).toBe(23);
  });

  it('returns nothing for anything that is not a date', () => {
    expect(parseFormDate('')).toBeNull();
    expect(parseFormDate('23/08/2026')).toBeNull();
    expect(parseFormDate(undefined)).toBeNull();
  });
});

describe('formatting helpers', () => {
  it('defaults a missing quantity to one', () => {
    expect(formatItems([{ item: 'Laptop' }])).toBe('Laptop x 1');
    expect(formatItems([{ item: 'Laptop', quantity: 'lots' }])).toBe('Laptop x 1');
  });

  it('rounds a fractional quantity down', () => {
    expect(formatItems([{ item: 'Laptop', quantity: 2.7 }])).toBe('Laptop x 2');
  });

  it('is calm about nothing at all', () => {
    expect(formatItems(undefined)).toBe('');
    expect(formatChecked(undefined)).toBe('');
  });
});

describe('signatureFileName', () => {
  it('strips the characters SharePoint refuses', () => {
    const name = signatureFileName(complete({ employeeName: 'A/B:C*D' }), 0);
    expect(name).not.toMatch(/[/\\?%*:|"<>]/);
    expect(name.endsWith('.png')).toBe(true);
  });
});

describe('todayValue', () => {
  /** Near midnight the local day and the UTC day differ. */
  it('is the local day, not the UTC one', () => {
    const now = new Date(2026, 7, 23, 23, 30);
    expect(todayValue(now)).toBe('2026-08-23');
  });
});

describe('requestForm', () => {
  it('labels the date by what the request is', () => {
    expect(dateLabel('Onboarding')).toBe('Join Date');
    expect(dateLabel('onboarding')).toBe('Join Date');
    expect(dateLabel('Offboarding')).toBe('Last Working Date');
    expect(dateLabel(undefined)).toBe('Last Working Date');
  });

  it('starts an employee with today and nothing else', () => {
    const employee = newEmployee();

    expect(employee.fullName).toBe('');
    expect(employee.equipmentItems).toEqual([]);
    expect(employee.joinDate).toMatch(/^\d{4}-\d{2}-\d{2}$/);
  });

  it('reads a SharePoint record back into the form', () => {
    const employee = employeeFromItem({
      Title: 'Amir Hakim',
      Calling_x0020_Name: 'Amir',
      Position: 'Engineer',
      Entity: 'PMW',
      Department: 'Engineering',
      Employee_x0020_ID: 'E-1042',
      Join_x0020__x002f__x0020_Last_x0: '2026-09-01T00:00:00Z',
      Equipment_x0020_Items: { results: ['laptop', 'monitor'] },
      Software_x0020_Licenses: { results: ['m365'] },
    });

    expect(employee.fullName).toBe('Amir Hakim');
    expect(employee.joinDate).toBe('2026-09-01');
    expect(employee.equipmentItems).toEqual(['laptop', 'monitor']);
    expect(employee.softwareLicenses).toEqual(['m365']);
  });

  /** A record with no multi-choice values must not read back as undefined. */
  it('reads absent multi-choice columns as empty lists', () => {
    const employee = employeeFromItem({ Title: 'Amir' });

    expect(employee.equipmentItems).toEqual([]);
    expect(employee.softwareLicenses).toEqual([]);
    expect(employee.joinDate).toBe('');
  });

  it('writes an employee out as the columns SharePoint expects', () => {
    const item = employeeToItem(newEmployee({
      fullName: 'Amir Hakim',
      entity: 'PMW',
      joinDate: '2026-09-01',
      equipmentItems: ['laptop'],
    }), 'Onboarding');

    expect(item.Title).toBe('Amir Hakim');
    expect(item.Request_x0020_Type).toBe('Onboarding');
    expect(item.Equipment_x0020_Items).toEqual(['laptop']);
  });

  /** SharePoint rejects an empty multi-choice list. */
  it('omits a multi-choice column with nothing chosen', () => {
    const item = employeeToItem(newEmployee(), 'Onboarding');

    expect('Equipment_x0020_Items' in item).toBe(false);
    expect('Software_x0020_Licenses' in item).toBe(false);
  });

  it('omits a date that is not one', () => {
    expect('Join_x0020__x002f__x0020_Last_x0' in employeeToItem(newEmployee({ joinDate: '' }), 'X'))
      .toBe(false);
  });

  /**
   * `new Date('2026-09-01')` is UTC midnight, which in Malaysia is 31 August at
   * 8am — a start date would store as the day before it was picked.
   */
  it('stores a picked day as that day', () => {
    const parsed = new Date(parseDay('2026-09-01'));

    expect(parsed.getMonth()).toBe(8);
    expect(parsed.getDate()).toBe(1);
  });

  it('allows up to ten employees on one submission, as before', () => {
    expect(MAX_EMPLOYEES).toBe(10);
  });
});

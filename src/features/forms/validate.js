import { fieldsFor, isRequest } from './checklistForm.js';

/**
 * What is wrong with a form, in words somebody can act on.
 *
 * Returned as a map of field to message rather than a boolean, so a step can
 * show everything that is missing at once. Being told about one empty field,
 * filling it, and being told about the next is the single most irritating way
 * a form can behave.
 */

const blank = (value) => !String(value ?? '').trim();

const LABELS = {
  formMode: 'Form type',
  employeeName: 'Employee name',
  employeeNo: 'Employee no',
  position: 'Position',
  entity: 'Entity',
  formDate: 'Date',
  signature: 'Signature',
};

const REQUIRED = ['employeeName', 'employeeNo', 'position', 'entity', 'formDate'];

export function validateChecklist(values, { step } = {}) {
  const errors = {};

  // Step 1 asks one question, and asking about step 2's fields while somebody
  // is still on step 1 would mark the whole form red before they had a chance.
  if (blank(values.formMode)) {
    errors.formMode = `${LABELS.formMode} is required.`;
    if (step === 0) return errors;
  }
  if (step === 0) return errors;

  const shown = new Set(fieldsFor(values.formMode));

  for (const field of REQUIRED) {
    if (shown.has(field) && blank(values[field])) {
      errors[field] = `${LABELS[field]} is required.`;
    }
  }

  if (isRequest(values.formMode)) {
    // At least one line with an item on it. A quantity without an item is not
    // a request for anything.
    const named = (values.items ?? []).filter((row) => !blank(row.item));
    if (!named.length) {
      errors.items = 'Add at least one item.';
    } else if (named.some((row) => !(Number(row.quantity) > 0))) {
      errors.items = 'Every item needs a quantity of at least 1.';
    }
  }

  // The signature is the entire point of this form: it is what makes it a
  // record of somebody accepting something rather than a note.
  if (!values.signature) errors.signature = 'Please sign before submitting.';

  return errors;
}

export const hasErrors = (errors) => Object.keys(errors).length > 0;

/**
 * The request form validates per EMPLOYEE, because one submission can carry
 * several. Errors come back keyed by index so the right panel can be marked
 * rather than the whole form.
 */
const EMPLOYEE_REQUIRED = {
  fullName: 'Full name',
  position: 'Position',
  entity: 'Entity',
  department: 'Department',
  joinDate: 'Date',
};

export function validateRequest(employees) {
  const errors = {};

  (employees ?? []).forEach((employee, index) => {
    const found = {};
    for (const [field, label] of Object.entries(EMPLOYEE_REQUIRED)) {
      if (blank(employee[field])) found[field] = `${label} is required.`;
    }
    if (Object.keys(found).length) errors[index] = found;
  });

  if (!employees?.length) errors.form = 'Add at least one employee.';

  return errors;
}

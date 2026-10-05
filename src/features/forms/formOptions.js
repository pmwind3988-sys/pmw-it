import { ENTITIES } from './checklistForm.js';
import { companyOptions, departmentOptions } from './orgDirectory.js';

/**
 * The Entity and Department choices a checklist offers, in one shape that
 * travels.
 *
 * The portal reads HR's `Companies` and `Departments` lists live. A shared link
 * cannot: the person opening it has no sign-in, and the server's identity is
 * granted the IT helpdesk site only, not HR's. So the portal takes a SNAPSHOT
 * when IT creates the link, the link carries it, and the public page and the
 * server both read the choices out of it. The portal page reads its live
 * directory through the same snapshot, so there is one set of rules.
 *
 * `options` is `{ entities: [{ value, label }], departments: { [entity]: [...] } }`,
 * or null when HR's lists could not be read — then the entity is one of the
 * built-in codes and the department is typed.
 */

export function snapshotOptions(directory) {
  if (!directory) return null;
  const entities = companyOptions(directory.companies);
  if (!entities.length) return null;

  const departments = {};
  for (const entity of entities) {
    departments[entity.value] = departmentOptions(directory.departments, entity.value);
  }
  return { entities, departments };
}

const asOption = (value) => ({ value, label: value });

export function entityChoices(options) {
  return options?.entities?.length ? options.entities : ENTITIES.map(asOption);
}

/**
 * `[]` — pick the entity first. `null` — type it: either HR's lists were not
 * available, or this company has no departments listed and a dropdown with
 * nothing in it could never be completed.
 */
export function departmentChoices(options, entity) {
  if (!options) return null;
  if (!entity) return [];
  const list = options.departments?.[entity];
  return list?.length ? list : null;
}

/** Does the department still belong to the entity? A stale pick is cleared. */
export function departmentFits(options, entity, department) {
  const choices = departmentChoices(options, entity);
  if (choices === null) return true;
  return choices.some((choice) => choice.value === department);
}

/** What to SHOW for a stored code: HR's name where there is one. */
export function optionLabel(options, values, field) {
  const value = values?.[field];
  if (!value) return value;
  const list = field === 'entity'
    ? entityChoices(options)
    : field === 'department' ? departmentChoices(options, values.entity) : null;
  return list?.find((choice) => choice.value === value)?.label ?? value;
}

/**
 * The values after an entity is picked. A department belongs to a company,
 * so one the new entity does not have is dropped rather than left on screen
 * to be signed for under the wrong company.
 */
export function withEntity(values, entity, options) {
  const next = { ...values, entity };
  if (next.department && !departmentFits(options, entity, next.department)) next.department = '';
  return next;
}

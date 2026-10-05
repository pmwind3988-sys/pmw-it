/**
 * The company and department lists the HR forms portal (`pmw-hrform`) keeps,
 * turned into dropdown options for the asset checklist.
 *
 * HR owns those two SharePoint lists, `Companies` and `Departments`, so that
 * every form says the same thing. This file only READS them — the rules below
 * are copied from `pmw-hrform/src/utils/orgDirectory.ts` and must stay the same
 * or the two portals will offer different departments for one company.
 *
 * A form stores the CODE and shows the NAME. A department's company is
 * optional: blank means every company, and a row that names the company beats
 * a shared row with the same code.
 *
 * Pure: no SharePoint, no React. `sharepoint/loadOrgDirectory.js` does the I/O.
 */

export const ORG_SITE_URL =
  import.meta.env?.VITE_HR_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/PMWHRDocs';

export const COMPANY_LIST = 'Companies';
export const DEPARTMENT_LIST = 'Departments';

const text = (value) => {
  if (typeof value === 'string') return value.trim();
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  return '';
};

/** Case- and spacing-insensitive, so "PMW  Lighting" and "pmw lighting" match. */
export const orgKey = (value) => text(value).toLowerCase().replace(/\s+/g, ' ');

/**
 * A blank or absent IsActive reads as ACTIVE. A column nobody has filled in
 * must never silently switch a company off.
 */
const activeFlag = (value) => (value === undefined || value === null || value === ''
  ? true
  : Boolean(value));

export function companyRows(items) {
  return (items ?? []).map((item) => ({
    name: text(item.Title),
    code: text(item.Code),
    isActive: activeFlag(item.IsActive),
  }));
}

export function departmentRows(items) {
  return (items ?? []).map((item) => ({
    name: text(item.Title),
    code: text(item.Code),
    company: text(item.Company),
    isActive: activeFlag(item.IsActive),
  }));
}

const byText = (a, b) => a.label.localeCompare(b.label);

/** Active companies as `{ value, label }`, the shape `SelectInput` reads. */
export function companyOptions(companies) {
  return (companies ?? [])
    .filter((company) => company.isActive && company.code)
    .map((company) => ({ value: company.code, label: company.name || company.code }))
    .sort(byText);
}

/**
 * The departments on offer once a company is picked.
 *
 * With no company picked yet, nothing is offered: the checklist asks for the
 * entity first, and a department list that ignores it would let somebody sign
 * for a department their company does not have.
 */
export function departmentOptions(departments, companyCode) {
  const wanted = orgKey(companyCode);
  if (!wanted) return [];

  const shared = new Map();
  const specific = new Map();

  for (const department of departments ?? []) {
    if (!department.isActive || !department.code) continue;
    const scope = orgKey(department.company);
    const key = orgKey(department.code);
    if (!scope) {
      if (!shared.has(key)) shared.set(key, department);
    } else if (scope === wanted && !specific.has(key)) {
      specific.set(key, department);
    }
  }

  // The specific row replaces the shared one of the same code.
  const merged = new Map([...shared, ...specific]);

  return [...merged.values()]
    .map((department) => ({
      value: department.code,
      label: department.name || department.code,
    }))
    .sort(byText);
}

/** Does this department still belong to the company? Used to clear a stale pick. */
export function departmentStillOffered(departments, companyCode, departmentCode) {
  const wanted = orgKey(departmentCode);
  return departmentOptions(departments, companyCode)
    .some((option) => orgKey(option.value) === wanted);
}

import { ENTITIES } from './checklistForm.js';

/**
 * The onboarding / offboarding request — HR or a manager raising an event, not
 * an employee signing for kit. It feeds `IT Request Form`, which the dashboard
 * and `/requests` are built on.
 *
 * Its fields are exactly what the SurveyJS version asked. Only the engine
 * underneath changed.
 */

export { ENTITIES };

/**
 * Dropdown options are read LIVE from the SharePoint columns rather than
 * declared here, because somebody adding a department in SharePoint expects the
 * form to offer it without a deploy. These are the columns to read.
 */
export const CHOICE_COLUMNS = [
  'Entity',
  'Equipment_x0020_Items',
  'Software_x0020_Licenses',
  'Request_x0020_Type',
  'Department',
];

export function newEmployee(defaults = {}) {
  return {
    fullName: '',
    callingName: '',
    position: '',
    entity: '',
    department: '',
    employeeId: '',
    joinDate: todayValue(),
    equipmentItems: [],
    equipmentRemarks: '',
    softwareLicenses: [],
    specialPermission: '',
    ...defaults,
  };
}

/** `yyyy-mm-dd` in the local day — near midnight the UTC one is yesterday. */
export function todayValue(now = new Date()) {
  const offset = now.getTimezoneOffset() * 60000;
  return new Date(now.getTime() - offset).toISOString().slice(0, 10);
}

/**
 * The same field means different things by request type: on an onboarding it is
 * when they start, on an offboarding it is their last day. One column holds
 * both, so only the label moves.
 */
export function dateLabel(requestType) {
  return String(requestType ?? '').toLowerCase() === 'onboarding'
    ? 'Join Date'
    : 'Last Working Date';
}

/** How many employees one submission may carry, as the old form allowed. */
export const MAX_EMPLOYEES = 10;

/** A SharePoint record back into the shape the form edits. */
export function employeeFromItem(item) {
  return newEmployee({
    fullName: item?.Title ?? '',
    callingName: item?.Calling_x0020_Name ?? '',
    position: item?.Position ?? '',
    entity: item?.Entity ?? '',
    department: item?.Department ?? '',
    employeeId: item?.Employee_x0020_ID ?? '',
    // A SharePoint datetime is an ISO instant; the date input wants the day.
    joinDate: item?.Join_x0020__x002f__x0020_Last_x0
      ? String(item.Join_x0020__x002f__x0020_Last_x0).split('T')[0]
      : '',
    equipmentItems: item?.Equipment_x0020_Items?.results ?? [],
    equipmentRemarks: item?.Equipment_x0020_Remarks ?? '',
    softwareLicenses: item?.Software_x0020_Licenses?.results ?? [],
    specialPermission: item?.Special_x0020_Permission ?? '',
  });
}

/**
 * One employee as the columns SharePoint expects. Multi-choice columns are
 * omitted when empty rather than sent as an empty list, which the REST endpoint
 * rejects.
 */
export function employeeToItem(employee, requestType) {
  const item = {
    Title: employee.fullName || '',
    Calling_x0020_Name: employee.callingName || '',
    Position: employee.position || '',
    Entity: employee.entity || '',
    Department: employee.department || '',
    Employee_x0020_ID: employee.employeeId || '',
    Equipment_x0020_Remarks: employee.equipmentRemarks || '',
    Special_x0020_Permission: employee.specialPermission || '',
    Request_x0020_Type: requestType || '',
  };

  const day = parseDay(employee.joinDate);
  if (day !== null) item.Join_x0020__x002f__x0020_Last_x0 = new Date(day).toISOString();

  if (employee.equipmentItems?.length) item.Equipment_x0020_Items = employee.equipmentItems;
  if (employee.softwareLicenses?.length) item.Software_x0020_Licenses = employee.softwareLicenses;

  return item;
}

/**
 * Parsed at local noon. `new Date('2026-09-01')` is UTC midnight, which in
 * Malaysia is the previous day at 8am — so a start date would store as the day
 * before it was picked.
 */
export function parseDay(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value ?? '').trim());
  if (!match) return null;

  const [, year, month, dayOfMonth] = match;
  const parsed = new Date(Number(year), Number(month) - 1, Number(dayOfMonth), 12, 0, 0);
  return Number.isNaN(parsed.getTime()) ? null : parsed.getTime();
}

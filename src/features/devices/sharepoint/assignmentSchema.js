/**
 * One row per STINT: somebody had this machine from one date to another.
 *
 * This list is the history; the device row's Owner / Department / Location
 * are the present. A machine therefore never loses its record by changing
 * hands -- "Carmen's old laptop now used by Aisyah" is the next row here.
 *
 * `DeviceId`, not the computer name, ties a stint to its machine, so a rename
 * does not cut the history in two. `Title` carries the name as it was, so the
 * list reads in SharePoint without a join.
 */
export const ASSIGNMENT_LIST_NAME = 'IT Device Assignments';

export const END_REASONS = ['Reassigned', 'Replaced', 'To stash', 'To repair', 'Retired'];

const text = (StaticName, Title) => ({ StaticName, Title, kind: 'text' });
const note = (StaticName, Title) => ({ StaticName, Title, kind: 'note' });
const num = (StaticName, Title) => ({ StaticName, Title, kind: 'number' });
const bool = (StaticName, Title) => ({ StaticName, Title, kind: 'boolean' });
const date = (StaticName, Title) => ({ StaticName, Title, kind: 'datetime' });
const choice = (StaticName, Title, choices) => ({ StaticName, Title, kind: 'choice', choices });

export const ASSIGNMENT_COLUMNS = [
  num('DeviceId', 'Device ID'),
  text('Owner', 'Owner'),
  text('Location', 'Location'),
  text('Department', 'Department'),
  date('AssignedOn', 'Assigned On'),
  bool('AssignedOnApprox', 'Start Is Approximate'),
  date('EndedOn', 'Ended On'),
  choice('EndReason', 'End Reason', END_REASONS),
  note('Note', 'Note'),
  text('RecordedBy', 'Recorded By'),
];

const iso = (ms) => (typeof ms === 'number' && Number.isFinite(ms) ? new Date(ms).toISOString() : null);
const msFromIso = (value) => (value ? new Date(value).getTime() : null);

export function toAssignmentItem(stint) {
  const item = {
    Title: stint.computerName ?? '',
    DeviceId: stint.deviceId,
    Owner: stint.owner ?? '',
    Location: stint.location ?? '',
    Department: stint.department ?? '',
    AssignedOnApprox: Boolean(stint.assignedOnApprox),
    Note: stint.note ?? '',
    RecordedBy: stint.recordedBy ?? '',
  };
  const assignedOn = iso(stint.assignedOn);
  if (assignedOn) item.AssignedOn = assignedOn;
  const endedOn = iso(stint.endedOn);
  if (endedOn) item.EndedOn = endedOn;
  if (stint.endReason) item.EndReason = stint.endReason;
  return item;
}

/** A partial write: only what closing a stint changes. */
export function closeBody({ endedOn, endReason, note: text }) {
  const body = { EndedOn: iso(endedOn), EndReason: endReason };
  if (text) body.Note = text;
  return body;
}

export function fromAssignmentItem(row) {
  const deviceId = row.DeviceId === null || row.DeviceId === undefined || row.DeviceId === ''
    ? null
    : Number(row.DeviceId);
  return {
    id: row.Id ?? row.ID ?? null,
    deviceId,
    computerName: row.Title || null,
    owner: row.Owner || null,
    location: row.Location || null,
    department: row.Department || null,
    assignedOn: msFromIso(row.AssignedOn),
    assignedOnApprox: row.AssignedOnApprox === true,
    endedOn: msFromIso(row.EndedOn),
    endReason: row.EndReason || null,
    note: row.Note || null,
    recordedBy: row.RecordedBy || null,
  };
}

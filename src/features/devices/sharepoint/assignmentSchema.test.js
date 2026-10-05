import { describe, it, expect } from 'vitest';
import {
  ASSIGNMENT_LIST_NAME, ASSIGNMENT_COLUMNS, END_REASONS,
  toAssignmentItem, fromAssignmentItem, closeBody,
} from './assignmentSchema.js';

const stint = {
  deviceId: 12, computerName: 'AMIR-HP', owner: 'Amir', location: 'F1', department: 'ENGINEERING',
  assignedOn: Date.UTC(2026, 1, 12, 4), assignedOnApprox: false,
  endedOn: null, endReason: null, note: null, recordedBy: 'it@pmw-group.com',
};

describe('assignment schema', () => {
  it('is its own list', () => expect(ASSIGNMENT_LIST_NAME).toBe('IT Device Assignments'));

  it('names exactly the five ways a stint ends', () => {
    expect(END_REASONS).toEqual(['Reassigned', 'Replaced', 'To stash', 'To repair', 'Retired']);
  });

  it('keeps the time of day on both dates', () => {
    for (const name of ['AssignedOn', 'EndedOn']) {
      expect(ASSIGNMENT_COLUMNS.find((c) => c.StaticName === name).kind).toBe('datetime');
    }
  });

  it('writes an open stint without an end', () => {
    const item = toAssignmentItem(stint);
    expect(item).toMatchObject({
      Title: 'AMIR-HP', DeviceId: 12, Owner: 'Amir', Location: 'F1', Department: 'ENGINEERING',
      AssignedOn: '2026-02-12T04:00:00.000Z', AssignedOnApprox: false, RecordedBy: 'it@pmw-group.com',
    });
    expect(item).not.toHaveProperty('EndedOn');
    expect(item).not.toHaveProperty('EndReason');
  });

  it('round-trips through SharePoint', () => {
    const back = fromAssignmentItem({ Id: 5, ...toAssignmentItem({ ...stint, endedOn: Date.UTC(2026, 9, 5), endReason: 'Reassigned' }) });
    expect(back).toMatchObject({ ...stint, id: 5, endedOn: Date.UTC(2026, 9, 5), endReason: 'Reassigned' });
  });

  it('closes a stint with only the closing columns', () => {
    expect(closeBody({ endedOn: Date.UTC(2026, 9, 5), endReason: 'To stash', note: null }))
      .toEqual({ EndedOn: '2026-10-05T00:00:00.000Z', EndReason: 'To stash' });
  });
});

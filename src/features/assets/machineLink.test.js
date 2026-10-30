import { describe, it, expect } from 'vitest';
import {
  machineFor, assetFor, machineMove, disagreements,
} from './machineLink.js';
import { ACTIONS } from '../devices/lifecycle/planLifecycle.js';

const laptop = {
  id: 7, category: 'Laptop', serialNumber: 'BVPZBC4', assetTag: 'PMWL010',
  status: 'Assigned', assignedTo: 'Nur Aisyah', assignedToEmail: 'aisyah@pmw.com',
};
const machine = { id: 41, computerName: 'PMWL010', serialNumber: 'bvpz bc4', status: 'Spare', location: 'F1' };

describe('machineFor', () => {
  it('links on the serial, whatever the spacing and case', () => {
    expect(machineFor(laptop, [machine])).toEqual({ device: machine, by: 'serial' });
  });

  it('links a computer on its label when the machine has no serial', () => {
    const hit = machineFor(laptop, [{ ...machine, serialNumber: 'Default string' }]);
    expect(hit?.by).toBe('name');
  });

  it('links when the label carries a second code beside the name', () => {
    const real = { ...laptop, serialNumber: 'BVPZBC4', assetTag: 'PMWL010 LTP-060526-3-SR-4' };
    expect(machineFor(real, [{ ...machine, serialNumber: '' }])?.by).toBe('name');
    expect(machineFor({ ...real, assetTag: 'PMWL0101' }, [{ ...machine, serialNumber: '' }])).toBeNull();
  });

  it('links on a name the machine had before a rename', () => {
    const renamed = { ...machine, serialNumber: '', computerName: 'AISYAH-PC', previousNames: ['PMWL010'] };
    expect(machineFor(laptop, [renamed])?.device).toBe(renamed);
  });

  it('never links a monitor on its label', () => {
    expect(machineFor({ ...laptop, category: 'Monitor', serialNumber: '' }, [{ ...machine, serialNumber: '' }])).toBeNull();
  });

  it('links nothing when a serial points at two machines', () => {
    expect(machineFor(laptop, [machine, { ...machine, id: 42 }])).toBeNull();
  });

  it('does not fall back to the label when the serial names a different machine', () => {
    const other = { id: 50, computerName: 'PMWL099', serialNumber: 'BVPZBC4' };
    expect(machineFor(laptop, [machine, other])).toBeNull();
    expect(machineFor(laptop, [other, { ...machine, serialNumber: 'XYZ1234' }])?.device).toBe(other);
  });
});

describe('assetFor', () => {
  it('finds the row from the machine', () => {
    expect(assetFor(machine, [laptop, { ...laptop, id: 8, serialNumber: 'OTHER', assetTag: 'PMWL011' }])).toBe(laptop);
  });

  it('links neither when two rows claim one machine', () => {
    expect(assetFor(machine, [laptop, { ...laptop, id: 8 }])).toBeNull();
  });
});

describe('machineMove', () => {
  it('assigns a stash machine to whoever the register says has it', () => {
    const move = machineMove(laptop, machine);
    expect(move.action).toBe(ACTIONS.ASSIGN);
    expect(move.expectedStatus).toBe('Spare');
    expect(move.input).toMatchObject({ owner: 'Nur Aisyah', location: 'F1' });
  });

  it('changes the owner when somebody else has it', () => {
    expect(machineMove(laptop, { ...machine, status: 'In use', owner: 'Ali' }).action).toBe(ACTIONS.CHANGE_OWNER);
  });

  it('does nothing when they already agree', () => {
    expect(machineMove(laptop, { ...machine, status: 'In use', owner: 'nur aisyah' })).toBeNull();
  });

  it('puts a machine back in the stash straight after a return', () => {
    const back = { ...laptop, status: 'In stock', assignedTo: '' };
    const used = { ...machine, status: 'In use', owner: 'Nur Aisyah' };
    expect(machineMove(back, used, { justReturned: true }).action).toBe(ACTIONS.TO_STASH);
    expect(machineMove(back, machine, { justReturned: true })).toBeNull();
  });

  it('does not call a counted, in-stock laptop that somebody uses a disagreement', () => {
    const counted = { ...laptop, status: 'In stock', assignedTo: '' };
    expect(machineMove(counted, { ...machine, status: 'In use', owner: 'Ali' })).toBeNull();
  });

  it('leaves repair and retirement in the register to the machine page', () => {
    expect(machineMove({ ...laptop, status: 'In repair' }, { ...machine, status: 'In use', owner: 'X' })).toBeNull();
  });
});

describe('disagreements', () => {
  it('lists each machine the register disagrees with, once', () => {
    const agreed = { ...laptop, id: 9, serialNumber: 'AAA111', assetTag: 'PMWL012' };
    const found = disagreements([laptop, agreed], [machine, { id: 43, serialNumber: 'AAA111', status: 'In use', owner: 'Nur Aisyah' }]);
    expect(found.map((entry) => entry.device.id)).toEqual([41]);
  });
});

import { describe, it, expect } from 'vitest';
import {
  summarise, worldTiles, locationTiles, machinesIn, searchMachines, PLACES,
} from './zones.js';
import { partGrades } from '../derive/partGrades.js';
import { defaultStandard } from '../standards/defaultStandard.js';

const S = defaultStandard();
const m = (over) => {
  const base = {
    id: Math.random(), computerName: 'PC', owner: 'A', location: 'F1', department: 'ENGINEERING',
    deviceType: 'Laptop', status: null, scanComplete: true,
    cpuGenerationRank: 12, installedRamGB: 32, storageTotalGB: 512, storageType: 'SSD only',
    dedicatedGpu: true, osSupported: true, windowsMajor: 11, ...over,
  };
  return { ...base, ...partGrades(base, S) };
};

describe('summarise', () => {
  it('counts laptops, desktops and anything else apart', () => {
    const s = summarise('F1', [m(), m({ deviceType: 'Desktop' }), m({ deviceType: 'Unknown' })]);
    expect(s).toMatchObject({ count: 3, laptops: 1, desktops: 1, other: 1 });
  });

  it('counts each part’s grades on its own, and the machines with a critical part', () => {
    const s = summarise('F1', [m({ installedRamGB: 4 }), m({ storageType: 'HDD only', installedRamGB: 4 }), m()]);
    expect(s.parts.ram.Critical).toBe(2);
    expect(s.parts.storage.Critical).toBe(1);
    expect(s.criticalMachines).toBe(2);
    expect(s.partBars.find((bar) => bar.key === 'ram').segs).toEqual([
      { grade: 'Critical', share: 2 / 3 }, { grade: 'Optimal', share: 1 / 3 },
    ]);
  });
});

describe('worldTiles', () => {
  const fleet = [
    m({ location: 'F1' }), m({ location: 'F1', department: 'SALES' }), m({ location: 'f1' }),
    m({ location: 'PML', department: 'FINANCE' }),
    m({ location: null }),
    m({ status: 'Spare', location: 'F1' }),
    m({ status: 'Retired', deviceType: 'Desktop' }),
  ];
  const world = worldTiles(fleet);

  it('has one tile per location, largest first, counting only machines in use or in repair', () => {
    expect(world.locations.map((t) => [t.code, t.count])).toEqual([['F1', 3], ['PML', 1]]);
  });

  it('lists a location\'s biggest departments as chips', () => {
    expect(world.locations[0].departments).toBe(2);
    expect(world.locations[0].chips[0]).toEqual({ name: 'ENGINEERING', count: 2 });
  });

  it('shows machines with no location only when there are some', () => {
    expect(world.noLocation.count).toBe(1);
    expect(worldTiles([m()]).noLocation).toBeNull();
  });

  it('keeps the stash and the graveyard apart from every location', () => {
    expect(world.stash.count).toBe(1);
    expect(world.graveyard).toMatchObject({ count: 1, desktops: 1 });
  });
});

describe('locationTiles', () => {
  it('keeps the same department at two locations apart', () => {
    const rows = [m({ location: 'F1', department: 'FINANCE' }), m({ location: 'PML', department: 'FINANCE' })];
    expect(locationTiles(rows, 'F1').departments.map((t) => t.count)).toEqual([1]);
  });

  it('puts machines with no department in Unassigned', () => {
    const tiles = locationTiles([m({ department: null }), m()], 'F1');
    expect(tiles.unassigned.count).toBe(1);
    expect(tiles.departments.map((t) => t.name)).toEqual(['ENGINEERING']);
    expect(tiles.centre.count).toBe(2);
  });
});

describe('machinesIn', () => {
  it('lists a department worst first', () => {
    const rows = [m({ computerName: 'B' }), m({ computerName: 'A', installedRamGB: 4 })];
    expect(machinesIn(rows, { location: 'F1', department: 'ENGINEERING' }).map((d) => d.computerName)).toEqual(['A', 'B']);
  });

  it('lists the stash, the graveyard and machines with no location', () => {
    const rows = [m({ status: 'Spare' }), m({ status: 'Retired' }), m({ location: '' })];
    expect(machinesIn(rows, { place: PLACES.STASH })).toHaveLength(1);
    expect(machinesIn(rows, { place: PLACES.GRAVEYARD })).toHaveLength(1);
    expect(machinesIn(rows, { place: PLACES.NO_LOCATION })).toHaveLength(1);
  });
});

describe('searchMachines', () => {
  it('finds by owner, computer name or serial', () => {
    const rows = [m({ owner: 'Amir', computerName: 'X1', serialNumber: '5CG1' }), m({ owner: 'Siti', computerName: 'X2' })];
    expect(searchMachines(rows, 'amir')).toHaveLength(1);
    expect(searchMachines(rows, '5cg')).toHaveLength(1);
    expect(searchMachines(rows, '  ')).toHaveLength(2);
  });
});

import { describe, it, expect } from 'vitest';
import { partGrades, cutoffGrade } from './partGrades.js';
import { defaultStandard } from '../standards/defaultStandard.js';

const S = defaultStandard();
const machine = (over) => ({
  department: 'ENGINEERING', scanComplete: true, deviceType: 'Laptop',
  cpuGenerationRank: 12, cpuAgeBand: 'Current', cpuModel: 'i7-1265U',
  installedRamGB: 32, storageTotalGB: 512, storageType: 'SSD only',
  dedicatedGpu: true, osSupported: true, windowsMajor: 11, windowsVersion: 'Microsoft Windows 11 Pro',
  ...over,
});

describe('cutoffGrade', () => {
  const c = { criticalBelow: 8, attentionBelow: 16, optimalFrom: 32 };
  it.each([[7, 'Critical'], [8, 'Needs attention'], [15, 'Needs attention'], [16, 'Moderate'], [31, 'Moderate'], [32, 'Optimal']])(
    '%i GB is %s', (value, grade) => expect(cutoffGrade(value, c)).toBe(grade),
  );
  it('skips an empty band', () => {
    expect(cutoffGrade(8, { criticalBelow: 8, attentionBelow: 8, optimalFrom: 16 })).toBe('Moderate');
  });
  it('has no grade for a missing value', () => expect(cutoffGrade(null, c)).toBeNull());
});

describe('partGrades', () => {
  it('grades a well-specified engineering machine Optimal on every part', () => {
    const g = partGrades(machine(), S);
    expect([g.gradeCpu, g.gradeRam, g.gradeStorage, g.gradeGraphics, g.gradeWindows]).toEqual(Array(5).fill('Optimal'));
    expect(g.actionRequired).toBe('Nothing to do');
    expect(g.personaKey).toBe('heavy');
  });

  it('grades each part on its own, with a reason naming the threshold', () => {
    const g = partGrades(machine({ installedRamGB: 16 }), S);
    expect(g.parts.ram).toEqual({ grade: 'Moderate', value: '16 GB', reason: 'Under the 32 GB needed for Optimal' });
    expect(g.gradeCpu).toBe('Optimal');
  });

  it('judges the same machine differently on a different desk', () => {
    expect(partGrades(machine({ installedRamGB: 16, department: 'FINANCE' }), S).gradeRam).toBe('Optimal');
  });

  it('makes an obsolete processor family Critical whatever the cut-offs', () => {
    const g = partGrades(machine({ cpuAgeBand: 'Obsolete', cpuGenerationRank: null, cpuModel: 'Celeron N4020' }), S);
    expect(g.parts.cpu.grade).toBe('Critical');
  });

  it('takes the worse of disk type and disk size', () => {
    expect(partGrades(machine({ storageTotalGB: 1000, storageType: 'HDD only' }), S).gradeStorage).toBe('Critical');
    expect(partGrades(machine({ storageTotalGB: 200, storageType: 'SSD only' }), S).gradeStorage).toBe('Needs attention');
  });

  it('asks a graphics card of engineering only', () => {
    expect(partGrades(machine({ dedicatedGpu: false }), S).gradeGraphics).toBe('Needs attention');
    expect(partGrades(machine({ dedicatedGpu: false, department: 'HR' }), S).gradeGraphics).toBe('Optimal');
  });

  it('places every Windows version', () => {
    const win10 = { windowsMajor: 10, windowsVersion: 'Microsoft Windows 10 Pro' };
    expect(partGrades(machine(win10), S).gradeWindows).toBe('Moderate');
    expect(partGrades(machine({ ...win10, osSupported: false }), S).gradeWindows).toBe('Critical');
    expect(partGrades(machine({ windowsMajor: 7, windowsVersion: 'Windows 7 Pro', osSupported: null }), S).gradeWindows).toBe('Critical');
    expect(partGrades(machine({ windowsMajor: null, windowsVersion: null, osSupported: null }), S).gradeWindows).toBe('Unknown');
  });

  it('never guesses a value the scan did not report', () => {
    const g = partGrades(machine({ cpuGenerationRank: null, installedRamGB: null, dedicatedGpu: null }), S);
    expect([g.gradeCpu, g.gradeRam, g.gradeGraphics]).toEqual(['Unknown', 'Unknown', 'Unknown']);
  });

  it('names the parts to act on', () => {
    const g = partGrades(machine({ installedRamGB: 4, storageType: 'HDD only', cpuGenerationRank: 9 }), S);
    expect(g.criticalParts).toEqual(['ram', 'storage']);
    expect(g.attentionParts).toEqual(['cpu']);
    expect(g.actionRequired).toBe('Upgrade now: RAM, Storage');
    expect(partGrades(machine({ cpuGenerationRank: 9 }), S).actionRequired).toBe('Plan: CPU');
  });

  it('grades nothing on an incomplete scan', () => {
    const g = partGrades(machine({ scanComplete: false }), S);
    expect(g.gradeRam).toBe('Unknown');
    expect(g.actionRequired).toBe('Re-run the scan');
  });

  it('uses the department map the standard carries', () => {
    const custom = defaultStandard();
    custom.departments.FINANCE = 'heavy';
    expect(partGrades(machine({ department: 'FINANCE', installedRamGB: 16 }), custom).gradeRam).toBe('Moderate');
  });
});

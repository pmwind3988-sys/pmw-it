import { describe, it, expect } from 'vitest';
import { fieldBody } from './provisionLists.js';

describe('fieldBody', () => {
  it('creates a text column as a plain SP.Field', () => {
    expect(fieldBody({ StaticName: 'Owner', Title: 'Owner', kind: 'text' })).toEqual({
      __metadata: { type: 'SP.Field' },
      Title: 'Owner', StaticName: 'Owner', FieldTypeKind: 2, Required: false,
    });
  });

  it('creates a DateTime column with the time kept', () => {
    const body = fieldBody({ StaticName: 'ScannedOn', Title: 'Scanned On', kind: 'datetime' });
    expect(body.__metadata.type).toBe('SP.FieldDateTime');
    expect(body.FieldTypeKind).toBe(4);
    // 1 = DateTime. 0 would be DateOnly and would throw the time away.
    expect(body.DisplayFormat).toBe(1);
  });

  it('creates a Note column as plain text, not rich text', () => {
    const body = fieldBody({ StaticName: 'RawReport', Title: 'Raw Report', kind: 'note' });
    expect(body.__metadata.type).toBe('SP.FieldMultiLineText');
    expect(body.FieldTypeKind).toBe(3);
    expect(body.RichText).toBe(false);
    expect(body.AppendOnly).toBe(false);
  });

  it('creates a choice column WITH its choices', () => {
    const body = fieldBody({
      StaticName: 'DeviceType', Title: 'Device Type', kind: 'choice',
      choices: ['Laptop', 'Desktop', 'Unknown'],
    });
    expect(body.FieldTypeKind).toBe(6);
    expect(body.Choices).toEqual({ results: ['Laptop', 'Desktop', 'Unknown'] });
  });

  it('creates a number column with no decimal places', () => {
    const body = fieldBody({ StaticName: 'InstalledRamGB', Title: 'RAM', kind: 'number' });
    expect(body.__metadata.type).toBe('SP.FieldNumber');
    expect(body.FieldTypeKind).toBe(9);
    expect(body.DisplayFormat).toBe(0);
  });

  it('creates a boolean column', () => {
    const body = fieldBody({ StaticName: 'HasHdd', Title: 'Has HDD', kind: 'boolean' });
    expect(body.FieldTypeKind).toBe(8);
  });
});

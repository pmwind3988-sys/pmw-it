import { describe, it, expect } from 'vitest';
import { mergeChoices, fieldBody, renameBody } from './provision.js';

describe('mergeChoices', () => {
  it('adds an option the column does not offer yet', () => {
    expect(mergeChoices(['pmw'], ['PMW', 'PCI'])).toEqual(['pmw', 'PMW', 'PCI']);
  });

  /**
   * The rule this function exists for. A SharePoint row saved with a value no
   * longer in its column's list becomes unreadable in that list, so dropping a
   * stale-looking option damages every record that used it.
   */
  it('never removes an option, however stale it looks', () => {
    const merged = mergeChoices(['pmw', 'pmw-ss', 'pmw-th'], ['PMW', 'PCI', 'PML', 'WB']);

    expect(merged).toContain('pmw-ss');
    expect(merged).toContain('pmw-th');
    expect(merged).toHaveLength(7);
  });

  it('does not duplicate an option that is already there', () => {
    expect(mergeChoices(['PMW', 'PCI'], ['PCI', 'PML'])).toEqual(['PMW', 'PCI', 'PML']);
  });

  it('leaves a column that already has everything unchanged', () => {
    const existing = ['PMW', 'PCI'];
    expect(mergeChoices(existing, ['PMW', 'PCI'])).toEqual(existing);
  });

  it('keeps the existing order, so a column does not reshuffle itself', () => {
    expect(mergeChoices(['b', 'a'], ['a', 'b', 'c'])).toEqual(['b', 'a', 'c']);
  });

  it('is calm about an empty declaration', () => {
    expect(mergeChoices(['PMW'], [])).toEqual(['PMW']);
  });
});

describe('fieldBody', () => {
  /**
   * `SP.Field` does not declare `Choices`; the tenant answers "The property
   * 'Choices' does not exist on type 'SP.Field'". This is the bug
   * `ensureAssetColumns` carried.
   */
  it('sends a choice column as SP.FieldChoice with its options', () => {
    const body = fieldBody({ StaticName: 'Entity', Title: 'Entity', kind: 'choice', choices: ['PMW'] });

    expect(body.__metadata.type).toBe('SP.FieldChoice');
    expect(body.Choices).toEqual({ results: ['PMW'] });
  });

  /** DisplayFormat 0 on a DateTime is DateOnly and silently discards the time. */
  it('keeps the time on a datetime column', () => {
    expect(fieldBody({ StaticName: 'A', Title: 'A', kind: 'datetime' }).DisplayFormat).toBe(1);
  });

  /** A rich-text Note wraps stored values in markup and will not round-trip. */
  it('creates a note as plain text', () => {
    expect(fieldBody({ StaticName: 'A', Title: 'A', kind: 'note' }).RichText).toBe(false);
  });

  /**
   * SharePoint derives the internal name from the Title a field is CREATED
   * with, so it is created under StaticName and renamed afterwards.
   */
  it('creates under the internal name, not the display name', () => {
    const column = { StaticName: 'FormDate', Title: 'Form Date', kind: 'datetime' };

    expect(fieldBody(column).Title).toBe('FormDate');
    expect(renameBody(column).Title).toBe('Form Date');
  });
});

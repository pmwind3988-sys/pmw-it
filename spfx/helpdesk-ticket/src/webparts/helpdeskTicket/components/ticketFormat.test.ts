import { formatBytes, friendlyUpdated, MAX_FILE_BYTES, splitBySize, tooBigMessage } from './ticketFormat';

describe('formatBytes', () => {
  it('reads sizes in B, KB and MB', () => {
    expect(formatBytes(512)).toBe('512 B');
    expect(formatBytes(2048)).toBe('2 KB');
    expect(formatBytes(3 * 1024 * 1024)).toBe('3.0 MB');
  });
});

describe('splitBySize', () => {
  const file = (name: string, size: number): File => ({ name, size } as File);

  it('refuses a file over 10 MB and keeps the rest', () => {
    const { ok, tooBig } = splitBySize([file('a.png', 1000), file('huge.mov', MAX_FILE_BYTES + 1), file('b.pdf', MAX_FILE_BYTES)]);
    expect(ok.map((f) => f.name)).toEqual(['a.png', 'b.pdf']);
    expect(tooBig.map((f) => f.name)).toEqual(['huge.mov']);
  });

  it('words the refusal for one file or several', () => {
    expect(tooBigMessage(['a.mov'])).toBe('a.mov is bigger than 10 MB, so it was not added.');
    expect(tooBigMessage(['a.mov', 'b.mov'])).toBe('a.mov, b.mov are bigger than 10 MB, so they were not added.');
    expect(tooBigMessage([])).toBe('');
  });
});

describe('friendlyUpdated', () => {
  const now = new Date(2026, 9, 13, 15, 0);

  it('says Today with the local time', () => {
    expect(friendlyUpdated(new Date(2026, 9, 13, 10, 40).toISOString(), now)).toBe('Today, 10:40 AM');
    expect(friendlyUpdated(new Date(2026, 9, 13, 0, 5).toISOString(), now)).toBe('Today, 12:05 AM');
  });

  it('says Yesterday, then the day and month', () => {
    expect(friendlyUpdated(new Date(2026, 9, 12, 9, 0).toISOString(), now)).toBe('Yesterday');
    expect(friendlyUpdated(new Date(2026, 9, 9, 9, 0).toISOString(), now)).toBe('9 Oct');
  });

  it('adds the year when it is another year', () => {
    expect(friendlyUpdated(new Date(2025, 9, 9, 9, 0).toISOString(), now)).toBe('9 Oct 2025');
  });

  it('is blank for a value that is not a date', () => {
    expect(friendlyUpdated('', now)).toBe('');
    expect(friendlyUpdated('not a date', now)).toBe('');
  });
});

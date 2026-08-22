// Parse + header detection + profiling, off the main thread -- spec §6.2.
//
// A 100k-row workbook takes seconds to parse and profile. On the main
// thread that is seconds of frozen UI with no way to show progress,
// because the thread that would paint the progress bar is the one doing
// the work. Here the main thread stays free to animate, and every stage
// reports before it starts rather than after it finishes.

import { parseWorkbook } from '../ingest/parseWorkbook.js';
import { detectHeader, toGrid } from '../ingest/detectHeader.js';
import { profileDataset } from '../profile/profileDataset.js';

self.onmessage = (e) => {
  const { type, arrayBuffer, sheetName, headerIndex: forcedHeaderIndex } = e.data ?? {};
  if (type !== 'parse') return;

  try {
    self.postMessage({ type: 'progress', stage: 'Reading workbook', pct: 10 });
    const { sheets } = parseWorkbook(arrayBuffer);
    if (!sheets.length) throw new Error('This workbook has no sheets.');

    const active = sheets.find((s) => s.name === sheetName) ?? sheets[0];

    self.postMessage({ type: 'progress', stage: 'Finding the header row', pct: 40 });
    // A forced index comes from the user correcting us in the profile
    // panel. Their choice is not re-scored -- they can see the sheet.
    const headerIndex = Number.isInteger(forcedHeaderIndex) && forcedHeaderIndex >= 0
      ? forcedHeaderIndex
      : detectHeader(active.rows).headerIndex;

    if (headerIndex === -1) {
      throw new Error(`No header row found in "${active.name}". Pick one manually.`);
    }
    const grid = toGrid(active.rows, headerIndex);

    self.postMessage({ type: 'progress', stage: 'Profiling columns', pct: 70 });
    const profile = profileDataset(grid);

    self.postMessage({
      type: 'parsed',
      sheets: sheets.map((s) => s.name),
      activeSheet: active.name,
      headerIndex,
      // The first rows of the raw sheet, so the profile screen can offer a
      // header-row picker showing what each candidate row actually says.
      headerCandidates: active.rows.slice(0, 20).map((row) => (row ?? []).map(
        (cell) => (cell instanceof Date ? cell.toISOString().slice(0, 10) : String(cell ?? '')),
      )),
      grid,
      profile,
    });
  } catch (err) {
    self.postMessage({ type: 'error', message: err?.message ?? 'Could not read that file.' });
  }
};

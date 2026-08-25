// Parse + header detection + profiling + cleaning, off the main thread
// -- spec §6.2.
//
// A 100k-row workbook takes seconds to parse and profile. On the main
// thread that is seconds of frozen UI with no way to show progress,
// because the thread that would paint the progress bar is the one doing
// the work. Here the main thread stays free to animate, and every stage
// reports before it starts rather than after it finishes.
//
// The parsed grid STAYS here after the parse. Every re-clean -- and the
// user re-cleans on every checkbox they tick -- then costs one small
// plan message in and typed arrays out, instead of shipping a
// hundred-thousand-row grid across the boundary each way.

import { parseWorkbook } from '../ingest/parseWorkbook.js';
import { detectHeader, toGrid } from '../ingest/detectHeader.js';
import { profileDataset } from '../profile/profileDataset.js';
import { applyCleanPlan } from '../clean/applyCleanPlan.js';

let currentGrid = null;

function handleParse({ arrayBuffer, sheetName, headerIndex: forcedHeaderIndex }) {
  self.postMessage({ type: 'progress', stage: 'Reading workbook', pct: 10 });
  const { sheets } = parseWorkbook(arrayBuffer);
  if (!sheets.length) throw new Error('This workbook has no sheets.');

  const active = sheets.find((s) => s.name === sheetName) ?? sheets[0];

  self.postMessage({ type: 'progress', stage: 'Finding the header row', pct: 40 });
  // A forced index comes from the user correcting us in the toolbar.
  // Their choice is not re-scored -- they can see the sheet.
  const headerIndex = Number.isInteger(forcedHeaderIndex) && forcedHeaderIndex >= 0
    ? forcedHeaderIndex
    : detectHeader(active.rows).headerIndex;

  if (headerIndex === -1) {
    throw new Error(`No header row found in "${active.name}". Pick one manually.`);
  }
  const grid = toGrid(active.rows, headerIndex);
  currentGrid = grid;

  self.postMessage({ type: 'progress', stage: 'Profiling columns', pct: 70 });
  const profile = profileDataset(grid);

  self.postMessage({
    type: 'parsed',
    sheets: sheets.map((s) => s.name),
    activeSheet: active.name,
    headerIndex,
    // The first rows of the raw sheet, so the toolbar's header-row
    // picker can show what each candidate row actually says.
    headerCandidates: active.rows.slice(0, 20).map((row) => (row ?? []).map(
      (cell) => (cell instanceof Date ? cell.toISOString().slice(0, 10) : String(cell ?? '')),
    )),
    grid,
    profile,
  });
}

function handleClean({ grid, profile, plan, requestId }) {
  // A grid on the message means the caller is reopening a dataset from
  // storage rather than continuing with the one just parsed. Adopting
  // it here keeps every later re-clean cheap in exactly the same way.
  if (grid) currentGrid = grid;
  if (!currentGrid) throw new Error('There is no imported sheet to clean.');
  const dataset = applyCleanPlan(currentGrid, plan, profile);
  // `requestId` lets the main thread ignore a result that a later toggle
  // has already superseded, so a slow clean cannot overwrite a fast one
  // that the user asked for afterwards.
  self.postMessage({ type: 'cleaned', dataset, requestId });
}

self.onmessage = (e) => {
  const msg = e.data ?? {};
  try {
    if (msg.type === 'parse') handleParse(msg);
    else if (msg.type === 'clean') handleClean(msg);
  } catch (err) {
    self.postMessage({
      type: 'error',
      message: err?.message ?? 'Could not read that file.',
      requestId: msg.requestId,
    });
  }
};

// The only two ways anything leaves this machine: a chart as a PNG and
// the answers behind it as a CSV.
//
// Nothing here is uploaded, and nothing is saved -- so these two ARE the
// sharing story. Both are cheap to get subtly wrong in ways that show up
// only in someone else's spreadsheet, which is why the CSV quoting and the
// date formatting are tested rather than eyeballed.

import { formatCell } from '../engine/formatCell.js';

/**
 * One CSV field, quoted only when it has to be.
 *
 * A field containing a comma, a quote or a newline must be wrapped in
 * quotes with its own quotes doubled. Skipping this is how one address
 * column silently shifts every later column of a row by one.
 */
export function csvField(value) {
  if (value === null || value === undefined) return '';
  const text = String(value);
  if (!/[",\r\n]/.test(text)) return text;
  return `"${text.replace(/"/g, '""')}"`;
}

/**
 * The cleaned dataset as CSV text.
 *
 * `mask` is optional: passing the canvas's current mask exports what the
 * user is looking at rather than everything, which is almost always what
 * "export this" means on a filtered screen.
 */
export function datasetToCsv(dataset, mask = null) {
  const lines = [dataset.columns.map((c) => csvField(c.name)).join(',')];

  for (let row = 0; row < dataset.rowCount; row++) {
    if (mask && !mask[row]) continue;
    lines.push(dataset.columns
      // A ratio stays a number here: formatted as "12.5%" it arrives in
      // Excel as text and no formula will touch it.
      .map((c) => csvField(formatCell(c, row, { percentAsText: false })))
      .join(','));
  }

  // CRLF, because Excel on Windows is the destination for essentially
  // every CSV this app produces.
  return lines.join('\r\n');
}

// --- browser side effects ---------------------------------------------

function download(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  // Revoked on the next tick rather than immediately: some browsers have
  // not started reading the blob by the time click() returns.
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

function safeName(name) {
  return String(name || 'export').replace(/[^a-z0-9\-_ ]+/gi, '').trim() || 'export';
}

/**
 * A tile as a PNG.
 *
 * `backgroundColor` is passed explicitly because ECharts' default is
 * transparent, and a transparent PNG pasted into a document or a chat
 * renders as dark-on-dark or light-on-light depending on where it lands.
 */
export function exportTilePng(chartInstance, title) {
  if (!chartInstance) return false;

  const panel = typeof document !== 'undefined'
    ? getComputedStyle(document.documentElement).getPropertyValue('--it-panel').trim()
    : '#ffffff';

  const url = chartInstance.getDataURL({
    type: 'png',
    pixelRatio: 2,
    backgroundColor: panel || '#ffffff',
  });

  const link = document.createElement('a');
  link.href = url;
  link.download = `${safeName(title)}.png`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  return true;
}

export function exportDatasetCsv(dataset, name, mask = null) {
  const csv = datasetToCsv(dataset, mask);
  // The BOM is what makes Excel read the file as UTF-8 rather than as
  // the local codepage, which is the difference between a name with an
  // accent in it and mojibake.
  // Written as an escape, never as a literal: a bare U+FEFF in source
  // is invisible in a diff and does not survive being retyped, which
  // would quietly drop the BOM and bring the mojibake back.
  const BOM = '\uFEFF';
  download(
    new Blob([`${BOM}${csv}`], { type: 'text/csv;charset=utf-8' }),
    `${safeName(name)}.csv`,
  );
}


// The styled .xlsx export. ExcelJS is about a megabyte, so it is fetched only
// when somebody presses Export (its own chunk), never with the page.

import { ExportColumn, safeText } from './exportColumns';
import { QueueTicket } from './ticket';

const HEADER_FILL = 'FF0B5CAD';
const HEADER_TEXT = 'FFFFFFFF';
const BAND_FILL = 'FFF3F8FD';
const BORDER = 'FFD0D7E2';

export interface ExportSummary {
  /** Label → value lines for the "About this export" sheet. */
  lines: [string, string][];
}

type ExcelModule = typeof import('exceljs');

async function loadExcel(): Promise<ExcelModule> {
  const mod = (await import(/* webpackChunkName: "exceljs" */ 'exceljs/dist/exceljs.min.js')) as unknown as
    { default?: ExcelModule } & ExcelModule;
  return mod.default ?? mod;
}

export async function toXlsx(rows: QueueTicket[], columns: ExportColumn[], summary: ExportSummary): Promise<Blob> {
  const Excel = await loadExcel();
  const wb = new Excel.Workbook();
  wb.creator = 'PMW IT helpdesk';
  wb.created = new Date();

  const ws = wb.addWorksheet('Tickets', { views: [{ state: 'frozen', ySplit: 1 }] });
  ws.columns = columns.map((c) => ({ header: c.header, key: c.key, width: c.width }));

  rows.forEach((t) => {
    const values: Record<string, unknown> = {};
    columns.forEach((c) => {
      const v = c.read(t);
      values[c.key] = typeof v === 'string' ? safeText(v) : v;
    });
    ws.addRow(values);
  });

  const header = ws.getRow(1);
  header.height = 22;
  header.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: HEADER_TEXT } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_FILL } };
    cell.alignment = { vertical: 'middle' };
    cell.border = { bottom: { style: 'thin', color: { argb: BORDER } } };
  });

  columns.forEach((c, i) => {
    const col = ws.getColumn(i + 1);
    col.alignment = { vertical: 'top', wrapText: !!c.wrap };
    if (rows.length && rows.some((t) => c.read(t) instanceof Date)) col.numFmt = 'dd mmm yyyy hh:mm';
  });

  // Every other row tinted, so a wide row can be followed across the sheet.
  for (let r = 2; r <= rows.length + 1; r++) {
    if (r % 2 === 1) {
      ws.getRow(r).eachCell({ includeEmpty: true }, (cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BAND_FILL } };
      });
    }
  }

  if (columns.length) {
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: Math.max(1, rows.length + 1), column: columns.length } };
  }

  const about = wb.addWorksheet('About this export');
  about.columns = [{ width: 22 }, { width: 70 }];
  summary.lines.forEach(([k, v]) => {
    const row = about.addRow([k, safeText(v)]);
    row.getCell(1).font = { bold: true };
    row.getCell(2).alignment = { wrapText: true, vertical: 'top' };
  });

  const buf = await wb.xlsx.writeBuffer();
  return new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

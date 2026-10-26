import { splitParts } from '../parse/parts.js';

/**
 * Which version of the scan script wrote the report, and whether it ran with
 * administrator rights. Version 2 added the licence and software blocks; the
 * `basic` variant is the fallback for a PC without the WMI tools, so it has
 * fewer of them filled in and must not overwrite a full scan.
 *
 * A report with no `Scan Info` block is a version 1 report.
 */
export function deriveScanInfo(fields) {
  const lines = fields['Scan Info'] ?? [];
  if (!lines.length) return { scanVersion: '1', scanRunAsAdmin: null };

  const cells = lines.map(splitParts).flat();
  const versionCell = cells.find((cell) => /^Version\s/i.test(cell));
  const adminCell = cells.find((cell) => /^Admin\s+(Yes|No)$/i.test(cell));

  // `Version 2` (PowerShell), `Version 2 exe` (the program) and `Version 2 basic`
  // (cmd only). The program reads exactly what PowerShell reads, so it is a full scan.
  let scanVersion = null;
  if (versionCell) {
    const raw = versionCell.replace(/^Version\s+/i, '').trim();
    const match = /^(\d+)(?:\s+(basic|exe))?$/i.exec(raw);
    if (match) scanVersion = /basic/i.test(match[2] ?? '') ? `${match[1]} basic` : match[1];
  }

  return {
    scanVersion,
    scanRunAsAdmin: adminCell ? /yes$/i.test(adminCell) : null,
  };
}

/**
 * Orders scan versions so a stored record can be told apart from an incoming
 * one: a full scan (2) outranks a basic one (2 basic), which outranks a v1
 * report. Null (a row from before the field existed) ranks lowest of all.
 */
export function scanRank(version) {
  if (version === null || version === undefined || version === '') return 0;
  if (/basic$/i.test(String(version))) return 2;
  const number = Number.parseInt(String(version), 10);
  if (!Number.isFinite(number)) return 0;
  // Version 1 is the old report; version 2 and later are full scans, which rank above basic.
  return number >= 2 ? number + 1 : number;
}

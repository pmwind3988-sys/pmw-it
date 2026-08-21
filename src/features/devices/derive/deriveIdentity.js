import { cleanValue } from '../parse/placeholders.js';
import { parsePairs, parseMailFiles } from '../parse/parseValues.js';

/** Longest first, so `PML GUARDHOUSE` matches before `PML` ever could. */
export const KNOWN_DEPARTMENTS = [
  'PML GUARDHOUSE', 'STOCKYARDF1', 'ENGINEERING', 'PRODUCTION', 'PURCHASING',
  'MARKETING', 'SHIPPING', 'FINANCE', 'ACCOUNT', 'SALES', 'ADMIN', 'STORE',
  'QAQC', 'QC', 'HR', 'IT',
];

// No trailing \b: `MS-7D99` has no word boundary between the 7 and the D, so
// `\bMS-7\b` would fail on the exact string this rule exists to catch.
const DESKTOP_BOARD = /(PRIME|MS-7\d|P5G|PRO B\d|TUF|ROG|\bH\d{3}M|\bB\d{3}M)/i;
const LAPTOP_MODEL =
  /Laptop|Notebook|Book|Pavilion|Inspiron|Latitude|Vostro|ThinkPad|IdeaPad|Precision\s+\d{4}|Folio|Elite/i;

const titleCase = (text) =>
  text
    .replace(/[._-]+/g, ' ')
    .trim()
    .split(/\s+/)
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(' ');

export function parseFileName(fileName) {
  const withoutExtension = fileName.replace(/\.txt$/i, '');
  const bracketMatch = /^\s*\[([^\]]+)\]\s*/.exec(withoutExtension);

  const bracket = bracketMatch ? bracketMatch[1].trim() : null;
  const rest = bracketMatch ? withoutExtension.slice(bracketMatch[0].length) : withoutExtension;

  return { bracket, stem: rest.replace(/_+$/, '').trim() };
}

function splitBracket(bracket) {
  if (!bracket) return { department: null, person: null };

  const upper = bracket.toUpperCase();
  const department = KNOWN_DEPARTMENTS.find(
    (dept) => upper === dept || upper.startsWith(`${dept} `),
  );

  if (!department) return { department: bracket, person: null };

  const remainder = bracket.slice(department.length).trim();
  return { department, person: remainder ? titleCase(remainder) : null };
}

function resolveOwner(fields, person) {
  const named = fields.Name?.length ? cleanValue(fields.Name[0]) : null;
  if (named) return { owner: named, ownerSource: 'Name field' };

  if (person) return { owner: person, ownerSource: 'Filename' };

  const credentials = parsePairs(fields['PMW Server and credentials'] ?? []);
  const username = credentials.find((pair) => pair.right)?.right;
  if (username) return { owner: titleCase(username), ownerSource: 'Server credential' };

  const mail = parseMailFiles(fields['Email data files found Active or Inactive account'] ?? []);
  // An .ost is the signed-in mailbox; a .pst is an archive that may belong to
  // somebody who left, so it is a weaker signal for "who uses this machine".
  const primary = mail.find((entry) => entry.kind === 'mailbox') ?? mail[0];
  if (primary?.file) {
    const localPart = primary.file.split('@')[0];
    if (localPart) return { owner: titleCase(localPart), ownerSource: 'Email' };
  }

  return { owner: null, ownerSource: null };
}

function resolveDeviceType(fields) {
  const model = fields['Computer Model']?.length ? fields['Computer Model'][0] : '';
  const board = fields.Motherboard?.length ? fields.Motherboard[0] : '';

  // The board is checked first because the computer NAME lies:
  // DESKTOP-2A3ERS8 is an HP EliteBook laptop.
  if (DESKTOP_BOARD.test(board)) return { deviceType: 'Desktop', deviceTypeConfident: true };
  if (LAPTOP_MODEL.test(model)) return { deviceType: 'Laptop', deviceTypeConfident: true };

  // An unset DMI product string means nobody flashed a model name in, which in
  // practice means a desktop assembled from parts.
  if (/^system product name$/i.test(model.trim())) {
    return { deviceType: 'Desktop', deviceTypeConfident: false };
  }

  return { deviceType: 'Unknown', deviceTypeConfident: false };
}

export function deriveIdentity(fields, fileName) {
  const { bracket, stem } = parseFileName(fileName);
  const { department, person } = splitBracket(bracket);

  const fromField = fields['Computer Name']?.length ? cleanValue(fields['Computer Name'][0]) : null;

  return {
    computerName: fromField ?? (stem || null),
    department,
    ...resolveOwner(fields, person),
    ...resolveDeviceType(fields),
  };
}

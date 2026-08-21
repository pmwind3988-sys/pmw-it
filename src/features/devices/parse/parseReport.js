import { KNOWN_LABELS, matchLabel } from './labels.js';

/** Lines the scan writes as decoration. They are not values. */
const SEPARATOR = /^=+$/;
const BANNERS = new Set(['COMPUTER INFORMATION', 'END OF REPORT']);

/**
 * A line that looks like `Something: value` but is not a known label. Used to
 * surface fields a future version of the scan script might add.
 */
const LABEL_SHAPED = /^([A-Za-z][\w /&()+.'-]{0,60}):\s*(.*)$/;

/** `C:\Users\...` and `Y: | \\server\...` are values, not labels. */
const DRIVE_LETTER = /^[A-Za-z]:/;

const INVISIBLE = /[\u00a0\u200b\u200c]/g;

function normaliseText(text) {
  return text
    .replace(/^\ufeff/, '')
    .replace(/\r\n?/g, '\n')
    .replace(INVISIBLE, ' ');
}

export function parseReport(text) {
  const fields = Object.fromEntries(KNOWN_LABELS.map((label) => [label, []]));
  const unknownLabels = [];
  const warnings = [];

  let current = null;
  // An unknown label owns the lines under it, exactly as a known one does.
  // Without this, `BitLocker Status:` followed by `Enabled` would file
  // "Enabled" under whichever known field came before it.
  let pendingUnknown = null;
  let sawKnownLabel = false;

  for (const raw of normaliseText(text).split('\n')) {
    const line = raw.replace(/\s+$/, '').trim();

    if (!line) continue;
    if (SEPARATOR.test(line) || BANNERS.has(line)) continue;

    const hit = matchLabel(line);
    if (hit) {
      sawKnownLabel = true;
      current = hit.label;
      pendingUnknown = null;
      if (hit.inline) fields[current].push(hit.inline);
      continue;
    }

    // Not a known label. Before treating it as a value, check whether it looks
    // like a label the scan script has newly started writing — but never treat
    // a pipe-delimited line or a drive path as one, because those are the two
    // shapes real values take in this format.
    const shaped = LABEL_SHAPED.exec(line);
    if (shaped && !line.includes(' | ') && !DRIVE_LETTER.test(line)) {
      pendingUnknown = { label: shaped[1].trim(), value: shaped[2].trim() };
      unknownLabels.push(pendingUnknown);
      current = null;
      continue;
    }

    if (pendingUnknown) {
      pendingUnknown.value = pendingUnknown.value ? `${pendingUnknown.value}\n${line}` : line;
    } else if (current) {
      fields[current].push(line);
    } else {
      warnings.push(`Value before any label: ${line}`);
    }
  }

  return { fields, unknownLabels, warnings, isReport: sawKnownLabel };
}

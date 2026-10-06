/**
 * What a department actually asks of a machine.
 *
 * The raw specification of a computer means nothing on its own: 8 GB and no
 * graphics card is a perfectly good guardhouse machine and a bad day for a
 * draughtsman. Every judgement in this section is therefore made against the
 * profile of the person sitting in front of it, not against one fleet-wide bar.
 */

/** Heavy: CAD, drawing, video, anything that pins a processor for minutes. */
const HEAVY = {
  key: 'heavy',
  label: 'Engineering / Technical / Media',
  blurb: 'Drawing, CAD or media work — needs cores, a real graphics card and headroom.',
  minRamGB: 8,
  goodRamGB: 16,
  comfortRamGB: 32,
  minCpuRank: 8,
  goodCpuRank: 10,
  needsDedicatedGpu: true,
  prefers: 'Desktop',
  mobility: 'low',
};

/** Standard desk work: mail, spreadsheets, the operations screens, a browser. */
const DESK = {
  key: 'desk',
  label: 'Logistics / Operations / Desk',
  blurb: 'Mail, spreadsheets and the operations screens — steady, unremarkable work.',
  minRamGB: 8,
  goodRamGB: 8,
  comfortRamGB: 16,
  minCpuRank: 6,
  goodCpuRank: 8,
  needsDedicatedGpu: false,
  prefers: 'Desktop',
  mobility: 'low',
};

/** People who work away from their desk more often than at it. */
const MOBILE = {
  key: 'mobile',
  label: 'Executive / Field',
  blurb: 'Works away from the desk — portability first, ordinary productivity specs.',
  minRamGB: 8,
  goodRamGB: 16,
  comfortRamGB: 16,
  minCpuRank: 7,
  goodCpuRank: 10,
  needsDedicatedGpu: false,
  prefers: 'Laptop',
  mobility: 'high',
};

/** A machine whose department the filename never told us. */
const UNKNOWN = {
  key: 'unknown',
  label: 'Unclassified',
  blurb: 'No department on the record, so it is judged against the plain desk baseline.',
  minRamGB: 8,
  goodRamGB: 8,
  comfortRamGB: 16,
  minCpuRank: 6,
  goodCpuRank: 8,
  needsDedicatedGpu: false,
  prefers: null,
  mobility: 'unknown',
};

export const PERSONAS = { HEAVY, DESK, MOBILE, UNKNOWN };

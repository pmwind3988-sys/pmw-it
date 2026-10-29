// Pure arithmetic for the dashboard's charts: where a point sits, how an axis is
// scaled, which labels can be shown without crowding, and how a share reads in
// words. Nothing here touches React or the page.

/** The coordinates every line and area plots in. The chart stretches to fit; these are only proportions. */
export const VIEW_W = 1000;
export const VIEW_H = 300;

/** Colours of the series, the bars and the heat ramp. Labels and numbers never take these. */
export const PALETTE = {
  raised: '#2a78d6',
  closed: '#eb6834',
  open: '#cde2fb',
  openLine: '#2a78d6',
  other: '#8a8886',
  calm: '#86b6ef',
  calmer: '#3987e5',
  deep: '#1c5cab',
  alert: '#ca5010',
  critical: '#a4262c',
  green: '#1baf7a',
  notSet: '#c8c6c4',
  empty: '#f3f2f1',
  hairline: '#edebe9',
  baseline: '#c8c6c4',
  heatRamp: ['#cde2fb', '#86b6ef', '#3987e5', '#1c5cab', '#0d366b']
};

export const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

const NICE_STEPS = [5, 10, 20, 25, 50, 100];

/** The smallest axis top that holds `value`, from 5, 10, 20, 25, 50, 100 and then the same times ten. */
export function niceMax(value: number): number {
  if (!(value > 0)) return 5;
  let scale = 1;
  for (let round = 0; round < 40; round += 1) {
    for (const step of NICE_STEPS) {
      if (step * scale >= value) return step * scale;
    }
    scale *= 10;
  }
  return value;
}

/** Tick values from 0 up to the axis top: four or five whole-number steps, so no label reads 1.25. */
export function axisTicks(max: number): number[] {
  const parts = max % 4 === 0 ? 4 : 5;
  const step = max / parts;
  const out: number[] = [];
  for (let i = 0; i <= parts; i += 1) out.push(i * step);
  return out;
}

/** Show every k-th label so that at most `most` labels appear along the axis. */
export function labelEvery(count: number, most = 12): number {
  if (count <= most) return 1;
  return Math.ceil(count / most);
}

/** "33%" of whole, rounded. Zero when there is nothing to take a share of. */
export function percentText(part: number, whole: number): string {
  if (!(whole > 0)) return '0%';
  return `${Math.round((part / whole) * 100)}%`;
}

/** A value as a percentage of the largest one, 0 to 100. */
export function barPct(value: number, max: number): number {
  if (!(max > 0) || !(value > 0)) return 0;
  return Math.min(100, (value / max) * 100);
}

/** Horizontal place of point i of n, in percent of the plot width. A single point sits in the middle. */
export function xPct(i: number, n: number): number {
  return n <= 1 ? 50 : (i / (n - 1)) * 100;
}

/** Vertical place of a value, in percent of the plot height from the top. Zero sits on the baseline. */
export function yPct(value: number, max: number): number {
  return 100 - barPct(value, max);
}

/** Vertical coordinate of a value inside the plot's viewBox. */
export function svgY(value: number, max: number): number {
  return (yPct(value, max) * VIEW_H) / 100;
}

const round1 = (x: number): number => Math.round(x * 10) / 10;

/** The "x,y x,y" list a polyline takes for a series of values. */
export function svgPoints(values: number[], max: number): string {
  const n = values.length;
  return values
    .map((v, i) => `${round1((xPct(i, n) * VIEW_W) / 100)},${round1(svgY(v, max))}`)
    .join(' ');
}

/** The closed outline of an area under a series, down to the baseline. */
export function svgArea(values: number[], max: number): string {
  return `M0,${VIEW_H} L${svgPoints(values, max)} L${VIEW_W},${VIEW_H} Z`;
}

/** A label is centred on its point, but kept inside the plot at either end. */
export function anchorShift(xp: number): string {
  if (xp <= 0) return '0%';
  if (xp >= 100) return '-100%';
  return '-50%';
}

/** Which point the pointer is nearest to, from its place across the chart (0 at the left edge, 1 at the right). */
export function snapIndex(ratio: number, count: number): number {
  if (count <= 1 || !(ratio > 0)) return 0;
  const r = Math.min(1, ratio);
  return Math.round(r * (count - 1));
}

/** One step left or right from the current point, never off either end. */
export function moveIndex(current: number | undefined, count: number, step: number): number {
  if (count <= 0) return 0;
  const base = current === undefined ? count - 1 : current;
  return Math.min(count - 1, Math.max(0, base + step));
}

/** A heat-map cell's colour: the empty colour for zero, else the ramp step its count reaches against the busiest cell. */
export function heatColour(count: number, max: number, ramp: string[], empty: string): string {
  if (!(count > 0) || !(max > 0) || ramp.length === 0) return empty;
  const index = Math.min(ramp.length - 1, Math.floor((count / max) * ramp.length));
  return ramp[index];
}

/** "1 ticket", "2 tickets". */
export function ticketsText(n: number): string {
  return `${n} ${n === 1 ? 'ticket' : 'tickets'}`;
}

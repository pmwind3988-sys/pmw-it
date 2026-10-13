import { GRADE_SLUG } from '../standards/gradeColors.js';
import { GRADES, UNKNOWN } from '../standards/defaultStandard.js';

/**
 * Geometry of the places map, in percent of its board. Nothing here touches
 * the DOM: the board positions its nodes with these numbers, the arrow keys
 * pick the next node from them, and the tests check them for overlaps.
 *
 * The reference board is 720 x 640 px on a desktop and 344 x 560 px on a phone;
 * percentages stretch with the board, so the px sizes are only used to judge
 * overlap, never to draw.
 */
export const HUB = { x: 50, y: 48 };
export const BOARD = { width: 720, height: 640 };
export const PHONE_BOARD = { width: 344, height: 560 };

/** Ring diameters in px, by the size class the CSS also uses. */
export const RING_PX = { lg: 160, md: 140, sm: 120 };

/** Fixed slots for the places that are not locations, as the mockup arranges them. */
export const EXTRA_SLOTS = {
  noLocation: { x: 50, y: 7 },
  stash: { x: 14, y: 86 },
  retired: { x: 50, y: 92 },
};
/** Rough footprint of each extra's shape in px, for the overlap test. */
export const EXTRA_PX = { noLocation: 74, stash: 118, retired: 76 };

const RX = 34;
const RY = 30;

/** Angles in degrees on the ellipse, 0 = right, -90 = top. Never exactly at top or bottom, where the extras sit. */
function ringAngles(count) {
  if (count <= 0) return [];
  if (count === 1) return [-150];
  if (count === 2) return [-150, -30];
  if (count === 3) return [-155, -25, 55];
  if (count === 4) return [-135, -45, 45, 135];
  const step = 360 / count;
  // A start of half a step keeps a node off the top and bottom when a whole number of steps lands on them.
  const start = Number.isInteger(90 / step) ? step / 2 : 0;
  return Array.from({ length: count }, (_, i) => start + i * step);
}

const round = (n) => Math.round(n * 100) / 100;

function onEllipse(angle) {
  const rad = (angle * Math.PI) / 180;
  return { x: round(HUB.x + RX * Math.cos(rad)), y: round(HUB.y + RY * Math.sin(rad)) };
}

/** Ring diameter class for a number of rings: bigger when there are few. */
export function ringSizeFor(count) {
  if (count <= 3) return 'lg';
  if (count <= 5) return 'md';
  return 'sm';
}

/**
 * Where every node of a level goes. `rings` is how many ring nodes (locations
 * or departments) sit round the hub; `extras` names any of the fixed places
 * (stash, noLocation, retired) that are shown. `ringSize` may be forced; by
 * default it shrinks as the count grows.
 */
export function mindMapLayout({ rings = 0, extras = [], ringSize } = {}) {
  const wanted = ringSize ?? ringSizeFor(rings);
  return {
    hub: { ...HUB },
    rings: ringAngles(rings).map(onEllipse),
    extras: Object.fromEntries(
      extras.filter((key) => EXTRA_SLOTS[key]).map((key) => [key, { ...EXTRA_SLOTS[key] }]),
    ),
    ringSize: wanted,
  };
}

/**
 * The curved branch from one point to another, as an SVG path in the 0-100
 * box. The control point sits to one side of the straight line, so every
 * branch bends the same way and the picture does not change between renders.
 */
export function curve(from, to) {
  const mx = (from.x + to.x) / 2;
  const my = (from.y + to.y) / 2;
  const dx = to.x - from.x;
  const dy = to.y - from.y;
  const cx = mx - dy * 0.18;
  const cy = my + dx * 0.18;
  return `M ${round(from.x)} ${round(from.y)} Q ${round(cx)} ${round(cy)} ${round(to.x)} ${round(to.y)}`;
}

const STEP = {
  ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0],
};

/**
 * The node an arrow key lands on: among the nodes that lie that way, the one
 * nearest in real distance, with a straight line across preferred. Distances
 * are taken in px on the board, so a wide board does not make sideways moves
 * look shorter than they are.
 */
export function nearestInDirection(points, index, key, board = BOARD) {
  const step = STEP[key];
  const from = points[index];
  if (!step || !from) return index;

  let best = index;
  let bestScore = Infinity;
  points.forEach((point, i) => {
    if (i === index) return;
    const dx = ((point.x - from.x) * board.width) / 100;
    const dy = ((point.y - from.y) * board.height) / 100;
    const along = dx * step[0] + dy * step[1];
    if (along <= 0) return;
    const across = Math.abs(step[0] ? dy : dx);
    const score = along + across * 2;
    if (score < bestScore) {
      bestScore = score;
      best = i;
    }
  });
  return best;
}

const DONUT_ORDER = [...GRADES, UNKNOWN];

/**
 * The conic gradient of a donut ring: one slice per worst-part grade, in
 * severity order, coloured from the grade variables so IT's standard colours
 * flow through. `worst` is the per-grade machine count from summarise().
 */
export function donutGradient(worst, total) {
  if (!total) return 'var(--grade-unknown)';
  const stops = [];
  let at = 0;
  const present = DONUT_ORDER.filter((grade) => (worst?.[grade] ?? 0) > 0);
  present.forEach((grade, i) => {
    const to = i === present.length - 1 ? 100 : at + ((worst[grade] / total) * 100);
    stops.push(`var(--grade-${GRADE_SLUG[grade]}) ${round(at)}% ${round(to)}%`);
    at = to;
  });
  return `conic-gradient(${stops.join(', ')})`;
}

/** Below this width, or above four rings, the circle has no room: the board becomes a tree. */
export const RING_MODE_MIN_WIDTH = 760;
export const RING_MODE_MAX_RINGS = 4;

/** Is the circle layout safe for this many rings on a board this wide? */
export const fitsCircle = (rings, width) => rings <= RING_MODE_MAX_RINGS && width >= RING_MODE_MIN_WIDTH;

const FLOW_EXTRA_ORDER = ['stash', 'noLocation', 'retired'];

/**
 * The tree layout: the hub on top, the rings in centred rows below it, the
 * extras in a row at the end. x is a percent of the board width and y is px
 * from its top; the board is as tall as `height`, so twelve departments make
 * a taller board rather than rings sitting on each other's captions. Each row
 * leaves a ring's diameter plus room for a three-line caption.
 * `nodes` are `{ key, kind }`; extras are recognised by kind.
 */
export function flowLayout(nodes, { width = 1000, ringPx = 140, phone = false } = {}) {
  const rings = nodes.filter((node) => !FLOW_EXTRA_ORDER.includes(node.kind));
  const extras = FLOW_EXTRA_ORDER.map((kind) => nodes.find((node) => node.kind === kind)).filter(Boolean);

  const colW = phone ? 150 : ringPx + 100;
  const cols = Math.max(phone ? 2 : 1, Math.floor((width - 16) / colW));
  const rowH = ringPx + (phone ? 150 : 140);
  const hubY = phone ? 64 : 90;
  const hubR = phone ? 48 : 75;
  const firstY = hubY + hubR + 56 + ringPx / 2;
  const step = (colW / Math.max(width, 1)) * 100;

  const at = {};
  rings.forEach((node, i) => {
    const row = Math.floor(i / cols);
    const inRow = Math.min(cols, rings.length - row * cols);
    const col = i % cols;
    at[node.key] = { x: round(50 + (col - (inRow - 1) / 2) * step), y: firstY + row * rowH };
  });

  const ringRows = Math.ceil(rings.length / cols);
  const lastRingY = rings.length ? firstY + (ringRows - 1) * rowH : hubY;
  const extrasY = rings.length ? lastRingY + rowH : firstY;
  const spread = phone ? { 1: [50], 2: [28, 72], 3: [18, 50, 82] } : { 1: [50], 2: [35, 65], 3: [25, 50, 75] };
  extras.forEach((node, i) => { at[node.key] = { x: spread[extras.length][i], y: extrasY }; });

  // A tree, not a starburst: the first row hangs off the hub and every later
  // node off the nearest node in the row above it, so one branch crosses each
  // gap instead of every branch running the whole height of the board.
  const rows = [];
  for (const node of [...rings, ...extras]) {
    const { y } = at[node.key];
    if (!rows.length || rows[rows.length - 1].y !== y) rows.push({ y, keys: [] });
    rows[rows.length - 1].keys.push(node.key);
  }
  const parent = {};
  rows.forEach((row, r) => {
    for (const key of row.keys) {
      if (r === 0) { parent[key] = null; continue; }
      const above = rows[r - 1].keys;
      parent[key] = above.reduce((best, k) => (
        Math.abs(at[k].x - at[key].x) < Math.abs(at[best].x - at[key].x) ? k : best), above[0]);
    }
  });

  const bottom = extras.length ? extrasY + 60 : lastRingY + ringPx / 2;
  // A last row of rings still has its captions below it; extras' captions are shorter.
  const pad = extras.length ? (phone ? 90 : 110) : (phone ? 170 : 150);
  return { hub: { x: 50, y: hubY }, at, parent, height: Math.round(bottom + pad) };
}

/**
 * A phone branch: down from the hub, then across to the node, as one smooth
 * S. `curve` bends sideways by a share of the vertical distance, which on a
 * phone board measured in px down and percent across would throw the branch
 * off the edge.
 */
export function phoneCurve(from, to) {
  const midY = round((from.y + to.y) / 2);
  return `M ${round(from.x)} ${round(from.y)} C ${round(from.x)} ${midY} ${round(to.x)} ${midY} ${round(to.x)} ${round(to.y)}`;
}

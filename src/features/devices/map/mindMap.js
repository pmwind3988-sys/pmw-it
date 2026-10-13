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

/** Phone geometry, in px down the board: the hub on top, rings zig-zagging below it, the extras in a row at the end. */
export const PHONE = { hubY: 64, firstRingY: 204, ringStep: 118, extrasGap: 150, bottomPad: 110 };
const PHONE_EXTRA_ORDER = ['stash', 'noLocation', 'retired'];
const PHONE_EXTRA_X = { 1: [50], 2: [28, 72], 3: [18, 50, 82] };

/**
 * Where everything goes on a phone, where the wide board's ellipse leaves no
 * room between a ring's caption and its neighbour. x is a percent of the board
 * width, y is px from its top, and the board is as tall as `height` -- so a
 * level with ten departments simply grows downward instead of overlapping.
 * `nodes` are `{ key, kind }`; extras are recognised by kind.
 */
export function phoneLayout(nodes) {
  const rings = nodes.filter((node) => !PHONE_EXTRA_ORDER.includes(node.kind));
  const extras = PHONE_EXTRA_ORDER
    .map((kind) => nodes.find((node) => node.kind === kind))
    .filter(Boolean);

  const at = {};
  rings.forEach((node, i) => {
    at[node.key] = { x: i % 2 ? 73 : 27, y: PHONE.firstRingY + i * PHONE.ringStep };
  });
  const lastRing = rings.length ? PHONE.firstRingY + (rings.length - 1) * PHONE.ringStep : PHONE.hubY;
  const extrasY = lastRing + PHONE.extrasGap;
  const xs = PHONE_EXTRA_X[extras.length] ?? [];
  extras.forEach((node, i) => { at[node.key] = { x: xs[i], y: extrasY }; });

  const height = (extras.length ? extrasY : lastRing) + PHONE.bottomPad;
  return { hub: { x: 50, y: PHONE.hubY }, at, height };
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

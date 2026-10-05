/**
 * Where each tile sits: a centre tile (the IT Stash, or the location itself),
 * the rest ringed round it in the order given -- callers pass largest first --
 * and on the world map the Graveyard below everything. Cells are 1-based CSS
 * grid lines in three equal columns, which is what lets the connector paths
 * be drawn in plain percentages with nothing measured.
 */
const RING = [
  { col: 1, row: 1 }, { col: 2, row: 1 }, { col: 3, row: 1 },
  { col: 1, row: 2 }, { col: 3, row: 2 },
  { col: 1, row: 3 }, { col: 3, row: 3 },
];

export function ringLayout(count, { reserveBottom = false } = {}) {
  const ring = reserveBottom ? RING : [...RING, { col: 2, row: 3 }];
  const cells = ring.slice(0, count);
  for (let row = 4; cells.length < count; row += 1) {
    for (let col = 1; col <= 3 && cells.length < count; col += 1) cells.push({ col, row });
  }

  const lowest = Math.max(3, ...cells.map((cell) => cell.row));
  const bottom = reserveBottom ? { col: 2, row: lowest > 3 ? lowest + 1 : 3 } : null;
  return {
    columns: 3,
    rows: bottom ? Math.max(lowest, bottom.row) : lowest,
    centre: { col: 2, row: 2 },
    cells,
    bottom,
  };
}

const STEP = {
  ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0],
};

/** The tile an arrow key lands on: the nearest one that way, straight lines preferred. */
export function neighbour(cells, index, key) {
  const step = STEP[key];
  const from = cells[index];
  if (!step || !from) return index;

  let best = index;
  let bestScore = Infinity;
  cells.forEach((cell, i) => {
    const dx = cell.col - from.col;
    const dy = cell.row - from.row;
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

const middle = (cell) => ({ x: (cell.col - 0.5) * 100, y: (cell.row - 0.5) * 100 });

export function connectors(centre, cells) {
  const from = middle(centre);
  return cells.map((cell) => {
    const to = middle(cell);
    return { x1: from.x, y1: from.y, x2: to.x, y2: to.y };
  });
}

import { useEffect, useRef, useState } from 'react';
import { Link } from 'react-router-dom';
import {
  HUB, RING_PX, curve, phoneCurve, nearestInDirection, flowLayout, fitsCircle,
} from '../map/mindMap';
import { GRADE_SLUG } from '../standards/gradeColors';

/** Which kind of branch leads to each kind of node: solid for places that hold machines, dashed for the stash and retired, dotted for no location. */
const BRANCH = {
  location: 'mm-branch-solid',
  dept: 'mm-branch-solid',
  unassigned: 'mm-branch-solid',
  stash: 'mm-branch-dashed',
  retired: 'mm-branch-dashed',
  noLocation: 'mm-branch-dotted',
};

/**
 * The mind map every level shares. The hub is the centre circle (never a link).
 * Each node is a link to its own page, drawn in one of four shapes: a donut
 * ring for a place's machines, a hexagon for the IT Stash, a diamond for no
 * location, a tombstone for retired. `nodes` are
 * `{ key, shape, kind, size, at, to, label, center, caption, donut }`, where `at`
 * is the position from mindMapLayout().
 *
 * Arrow keys move between the nodes by position; the handler sits on each link,
 * never on window, so it cannot fight a text box.
 *
 * Two layouts, chosen by the board's measured width: the circle (the hub in
 * the middle, rings round it) when a few rings fit, and the tree (the hub on
 * top, rings in rows) when they would not -- a place with twelve departments,
 * or any phone. The tree's board grows downward, so nothing ever overlaps.
 */
export default function MindMap({ label, hub, nodes }) {
  const refs = useRef([]);
  const board = useRef(null);
  const [width, setWidth] = useState(() => (typeof window === 'undefined' ? 1000 : Math.max(320, window.innerWidth - 48)));

  useEffect(() => {
    const element = board.current;
    if (!element || typeof ResizeObserver === 'undefined') return undefined;
    const observer = new ResizeObserver(([entry]) => setWidth(Math.round(entry.contentRect.width)));
    observer.observe(element);
    return () => observer.disconnect();
  }, []);

  const rings = nodes.filter((node) => node.shape === 'ring');
  const circle = fitsCircle(rings.length, width);
  const phone = width < 600;
  const ringPx = phone ? 104 : RING_PX[rings[0]?.size ?? 'md'];
  const flow = circle ? null : flowLayout(nodes, { width, ringPx, phone });

  // Circle: both axes in percent of a fixed board. Tree: x in percent, y in px of a board as tall as the layout.
  const hubAt = circle ? HUB : flow.hub;
  const atOf = (node) => (circle ? node.at : flow.at[node.key]);
  const place = (at) => ({ left: `${at.x}%`, top: circle ? `${at.y}%` : `${at.y}px` });
  const branch = (node) => {
    if (circle) return curve(HUB, node.at);
    const up = flow.parent[node.key];
    return phoneCurve(up ? flow.at[up] : flow.hub, flow.at[node.key]);
  };
  const points = nodes.map(atOf);
  const keyBoard = circle ? undefined : { width, height: 100 };

  const onKeyDown = (index) => (event) => {
    if (!event.key.startsWith('Arrow')) return;
    const next = nearestInDirection(points, index, event.key, keyBoard);
    if (next !== index && refs.current[next]) {
      event.preventDefault();
      refs.current[next].focus();
    }
  };

  return (
    <div className="mm">
      <div
        ref={board}
        className={`mm-board${circle ? '' : ' mm-board-tree'}${phone ? ' mm-board-compact' : ''}`}
        role="group"
        aria-label={label}
        style={circle ? undefined : { height: `${flow.height}px` }}
      >
        <svg className="mm-branches" viewBox={circle ? '0 0 100 100' : `0 0 100 ${flow.height}`} preserveAspectRatio="none" aria-hidden="true">
          {nodes.map((node) => (
            <path key={node.key} d={branch(node)} className={BRANCH[node.kind] ?? 'mm-branch-solid'} />
          ))}
        </svg>

        <div className="mm-node mm-hub" style={place(hubAt)}>
          <span className="mm-hole">{hub}</span>
        </div>

        {nodes.map((node, index) => (
          <Link
            key={node.key}
            to={node.to}
            aria-label={node.label}
            ref={(el) => { refs.current[index] = el; }}
            onKeyDown={onKeyDown(index)}
            className={`mm-node mm-${node.shape} mm-${node.kind}${node.size ? ` mm-${node.size}` : ''}${node.donut ? ' mm-graded' : ''}`}
            style={{ ...place(atOf(node)), ...(node.donut ? { background: node.donut } : null) }}
          >
            {node.shape === 'ring'
              ? <span className="mm-hole">{node.center}</span>
              : <span className="mm-shape">{node.center}</span>}
            <span className="mm-caption">{node.caption}</span>
          </Link>
        ))}
      </div>

      <p className="mm-keys"><kbd>← ↑ → ↓</kbd> move <kbd>Enter</kbd> go in <kbd>Esc</kbd> back out</p>
    </div>
  );
}

const LEGEND = [
  ['Critical', 'Critical'],
  ['Needs attention', 'Needs attention'],
  ['Moderate', 'Moderate'],
  ['Optimal', 'Optimal'],
];

/** The four grades as coloured dots, read straight from the standard's colour variables. */
export function GradeLegend() {
  return (
    <div className="mm-legend" aria-label="Legend: each ring is that place's machines, each counted once at its worst part">
      {LEGEND.map(([grade, text]) => (
        <span key={grade} className="mm-legend-item">
          <span className={`mm-dot g-${GRADE_SLUG[grade]}`} aria-hidden="true" />
          {text}
        </span>
      ))}
    </div>
  );
}

import { useRef } from 'react';
import { Link } from 'react-router-dom';
import {
  HUB, curve, phoneCurve, nearestInDirection, phoneLayout,
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
 * Every node carries two positions as CSS variables: the wide board's (percent
 * of both sides) and the phone's (phoneLayout: percent across, px down, on a
 * board that grows with the count). The stylesheet picks one by width, so the
 * same links serve both and nothing is rendered twice.
 */
export default function MindMap({ label, hub, nodes }) {
  const refs = useRef([]);
  const points = nodes.map((node) => node.at);
  const phone = phoneLayout(nodes);
  const place = (wide, narrow) => ({
    '--x': `${wide.x}%`, '--y': `${wide.y}%`, '--px': `${narrow.x}%`, '--py': `${narrow.y}px`,
  });

  const onKeyDown = (index) => (event) => {
    if (!event.key.startsWith('Arrow')) return;
    const next = nearestInDirection(points, index, event.key);
    if (next !== index && refs.current[next]) {
      event.preventDefault();
      refs.current[next].focus();
    }
  };

  return (
    <div className="mm">
      <div className="mm-board" role="group" aria-label={label} style={{ '--phone-h': `${phone.height}px` }}>
        <svg className="mm-branches mm-branches-wide" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
          {nodes.map((node) => (
            <path key={node.key} d={curve(HUB, node.at)} className={BRANCH[node.kind] ?? 'mm-branch-solid'} />
          ))}
        </svg>
        {/* The phone's branches live in a box exactly as tall as its board: x in percent, y in px. */}
        <svg className="mm-branches mm-branches-phone" viewBox={`0 0 100 ${phone.height}`} preserveAspectRatio="none" aria-hidden="true">
          {nodes.map((node) => (
            <path key={node.key} d={phoneCurve(phone.hub, phone.at[node.key])} className={BRANCH[node.kind] ?? 'mm-branch-solid'} />
          ))}
        </svg>

        <div className="mm-node mm-hub" style={place(HUB, phone.hub)}>
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
            style={{ ...place(node.at, phone.at[node.key]), ...(node.donut ? { background: node.donut } : null) }}
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

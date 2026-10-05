import { useRef } from 'react';
import { connectors, neighbour } from '../map/mapLayout';

/**
 * The board the tiles sit on. Arrow keys move between tiles by position;
 * the handler is on each tile, never on window, so it cannot fight a text box.
 * `items` = [{ key, cell, render(tileRef, onKeyDown) }], centre and bottom included.
 */
export default function MapGrid({ layout, items }) {
  const refs = useRef([]);
  const cells = items.map((item) => item.cell);
  const linkable = items.map((item) => item.linkable !== false);

  const onKeyDown = (index) => (event) => {
    if (!event.key.startsWith('Arrow')) return;
    let next = neighbour(cells, index, event.key);
    // Skip a tile that is not a link (a location's own centre tile).
    if (!linkable[next]) next = neighbour(cells, next, event.key);
    if (next !== index && refs.current[next]) {
      event.preventDefault();
      refs.current[next].focus();
    }
  };

  const paths = connectors(layout.centre, items.filter((item) => item.cell !== layout.centre).map((item) => item.cell));

  return (
    <div className="mg-wrap">
      <div className="mg" style={{ '--rows': layout.rows }}>
        <svg className="mg-paths" viewBox={`0 0 ${layout.columns * 100} ${layout.rows * 100}`} preserveAspectRatio="none" aria-hidden="true">
          {paths.map((p) => (
            <line key={`${p.x2},${p.y2}`} x1={p.x1} y1={p.y1} x2={p.x2} y2={p.y2} vectorEffect="non-scaling-stroke" />
          ))}
        </svg>
        {/* eslint-disable-next-line react-hooks/refs */}
        {items.map((item, index) => item.render((node) => { refs.current[index] = node; }, onKeyDown(index)))}
      </div>
    </div>
  );
}

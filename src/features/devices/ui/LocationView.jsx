import { useMemo } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import MindMap from './MindMap';
import { locationTiles } from '../map/zones';
import { mindMapLayout, donutGradient } from '../map/mindMap';
import { mapHref } from '../map/mapLinks';
import { profileKeyFor } from '../standards/defaultStandard.js';
import { plural } from '../../../utils/plural';

export default function LocationView({ devices, location, standard }) {
  const navigate = useNavigate();
  const tiles = useMemo(() => locationTiles(devices, location), [devices, location]);
  const ringed = [...tiles.departments, ...(tiles.unassigned ? [tiles.unassigned] : [])];
  const layout = mindMapLayout({ rings: ringed.length, ringSize: 'md' });

  const nodes = ringed.map((tile, i) => {
    const unassigned = tile.name === 'Unassigned';
    const profile = unassigned ? 'No department yet' : standard.profiles[profileKeyFor(standard, tile.name)].label;
    return {
      key: `dept-${tile.name}`,
      shape: 'ring',
      kind: unassigned ? 'unassigned' : 'dept',
      size: layout.ringSize,
      at: layout.rings[i],
      to: mapHref({ location, department: tile.name }),
      label: `${tile.name}, ${tile.count} machines, ${tile.workMachines} need work, ${profile}`,
      donut: donutGradient(tile.worst, tile.count),
      center: (<><strong className={`mm-name${tile.name.length > 7 ? ' mm-name-long' : ''}`}>{tile.name}</strong><span className="mm-sub">{plural(tile.count, 'machine')}</span></>),
      caption: (<><span className="mm-pill">{tile.workMachines} need work</span><span className="mm-line">{profile}</span><span className="mm-line">{plural(tile.laptops, 'laptop')} · {plural(tile.desktops, 'desktop')}</span></>),
    };
  });

  return (
    <section className="dm" onKeyDown={(event) => { if (event.key === 'Escape') navigate(mapHref()); }}>
      <nav className="dm-crumbs" aria-label="Breadcrumb">
        <Link to={mapHref()}>Map</Link><span aria-hidden="true">›</span><span aria-current="page">{tiles.centre.name}</span>
      </nav>
      <header className="dm-head">
        <span className="dm-eyebrow">Location · {tiles.centre.name}</span>
        <h2 className="dm-title">Choose a department</h2>
        <p className="dm-sub">
          {plural(tiles.departments.length, 'department')} · {plural(tiles.centre.count, 'machine')} · {plural(tiles.centre.laptops, 'laptop')} · {plural(tiles.centre.desktops, 'desktop')}
        </p>
      </header>
      {tiles.centre.count === 0
        ? <p className="dm-empty">Nothing is recorded at {tiles.centre.name} yet.</p>
        : (
          <MindMap
            label={`${tiles.centre.name} departments`}
            hub={(<><span className="mm-hub-eyebrow">{tiles.centre.name}</span><strong className="mm-hub-count">{tiles.centre.count}</strong><span className="mm-hub-sub">machines</span></>)}
            nodes={nodes}
          />
        )}
    </section>
  );
}

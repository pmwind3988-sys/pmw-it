import { useMemo } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import MapTile from './MapTile';
import MapGrid from './MapGrid';
import { locationTiles } from '../map/zones';
import { ringLayout } from '../map/mapLayout';
import { mapHref } from '../map/mapLinks';
import { personaFor } from '../derive/persona';

export default function LocationView({ devices, location }) {
  const navigate = useNavigate();
  const tiles = useMemo(() => locationTiles(devices, location), [devices, location]);
  const ringed = [...tiles.departments, ...(tiles.unassigned ? [tiles.unassigned] : [])];
  const layout = ringLayout(ringed.length);

  const items = [
    {
      key: 'centre',
      cell: layout.centre,
      linkable: false,
      render: () => (
        <MapTile key="centre" eyebrow="Location" title={tiles.centre.name} summary={tiles.centre} variant="hub" cell={layout.centre} />
      ),
    },
    ...ringed.map((tile, index) => ({
      key: tile.name,
      cell: layout.cells[index],
      render: (ref, onKeyDown) => (
        <MapTile key={tile.name} to={mapHref({ location, department: tile.name })}
          eyebrow={personaFor(tile.name === 'Unassigned' ? null : tile.name).label ?? 'Department'}
          title={tile.name} summary={tile} cell={layout.cells[index]} tileRef={ref} onKeyDown={onKeyDown} />
      ),
    })),
  ];

  return (
    <section className="dm" onKeyDown={(event) => { if (event.key === 'Escape') navigate(mapHref()); }}>
      <nav className="dm-crumbs" aria-label="Breadcrumb">
        <Link to={mapHref()}>Map</Link><span aria-hidden="true">›</span><span aria-current="page">{tiles.centre.name}</span>
      </nav>
      <header className="dm-head">
        <span className="dm-eyebrow">Location · {tiles.centre.name}</span>
        <h2 className="dm-title">Choose a department</h2>
        <p className="dm-sub">
          {tiles.departments.length} departments · {tiles.centre.count} machines · {tiles.centre.laptops} laptops · {tiles.centre.desktops} desktops
        </p>
      </header>
      {tiles.centre.count === 0
        ? <p className="dm-empty">Nothing is recorded at {tiles.centre.name} yet.</p>
        : <MapGrid layout={layout} items={items} />}
    </section>
  );
}

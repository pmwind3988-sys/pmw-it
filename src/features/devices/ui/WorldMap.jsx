import { useMemo } from 'react';
import MapTile from './MapTile';
import MapGrid from './MapGrid';
import { Archive, Tombstone } from '../../../components/ui/Icons';
import { worldTiles, PLACES } from '../map/zones';
import { ringLayout } from '../map/mapLayout';
import { mapHref } from '../map/mapLinks';

export default function WorldMap({ devices }) {
  const world = useMemo(() => worldTiles(devices), [devices]);
  const ringed = [
    ...world.locations.map((tile) => ({ tile, to: mapHref({ location: tile.code }), eyebrow: 'Location', variant: 'zone' })),
    ...(world.noLocation ? [{ tile: world.noLocation, to: mapHref({ place: PLACES.NO_LOCATION }), eyebrow: 'Needs a location', variant: 'pending' }] : []),
  ];
  const layout = ringLayout(ringed.length, { reserveBottom: true });

  // DOM order is the phone order: stash, locations, graveyard.
  const items = [
    {
      key: 'stash',
      cell: layout.centre,
      render: (ref, onKeyDown) => (
        <MapTile key="stash" to={mapHref({ place: PLACES.STASH })} eyebrow="Hub" title="IT Stash" summary={world.stash}
          variant="hub" cell={layout.centre} tileRef={ref} onKeyDown={onKeyDown}>
          <Archive size={20} className="mt-glyph" />
        </MapTile>
      ),
    },
    ...ringed.map((entry, index) => ({
      key: entry.tile.name,
      cell: layout.cells[index],
      render: (ref, onKeyDown) => (
        <MapTile key={entry.tile.name} to={entry.to} eyebrow={entry.eyebrow} title={entry.tile.name}
          summary={entry.tile} chips={entry.tile.chips} variant={entry.variant}
          cell={layout.cells[index]} tileRef={ref} onKeyDown={onKeyDown} />
      ),
    })),
    {
      key: 'graveyard',
      cell: layout.bottom,
      render: (ref, onKeyDown) => (
        <MapTile key="graveyard" to={mapHref({ place: PLACES.GRAVEYARD })} eyebrow="Graveyard" title="Retired"
          summary={world.graveyard} variant="grave" cell={layout.bottom} tileRef={ref} onKeyDown={onKeyDown}>
          <Tombstone size={18} className="mt-glyph" />
        </MapTile>
      ),
    },
  ];

  const fleet = world.locations.reduce((sum, t) => sum + t.count, 0) + (world.noLocation?.count ?? 0);
  const laptops = world.locations.reduce((sum, t) => sum + t.laptops, 0) + (world.noLocation?.laptops ?? 0);
  const desktops = world.locations.reduce((sum, t) => sum + t.desktops, 0) + (world.noLocation?.desktops ?? 0);

  return (
    <section className="dm">
      <header className="dm-head">
        <span className="dm-eyebrow">PMW fleet</span>
        <h2 className="dm-title">Choose a location</h2>
        <p className="dm-sub">{fleet} machines in use · {laptops} laptops · {desktops} desktops</p>
      </header>
      <MapGrid layout={layout} items={items} />
      <p className="dm-keys"><kbd>← ↑ → ↓</kbd> move <kbd>Enter</kbd> go in <kbd>Esc</kbd> back out</p>
    </section>
  );
}

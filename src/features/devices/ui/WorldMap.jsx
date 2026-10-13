import { useMemo } from 'react';
import MindMap, { GradeLegend } from './MindMap';
import { worldTiles, PLACES } from '../map/zones';
import { mindMapLayout, donutGradient } from '../map/mindMap';
import { mapHref } from '../map/mapLinks';
import { plural } from '../../../utils/plural';

export default function WorldMap({ devices }) {
  const world = useMemo(() => worldTiles(devices), [devices]);
  const extras = ['stash', ...(world.noLocation ? ['noLocation'] : []), 'retired'];
  const layout = mindMapLayout({ rings: world.locations.length, extras });

  const nodes = [
    ...world.locations.map((tile, i) => ({
      key: `loc-${tile.code}`,
      shape: 'ring',
      kind: 'location',
      size: layout.ringSize,
      at: layout.rings[i],
      to: mapHref({ location: tile.code }),
      label: `${tile.code}, ${tile.count} machines, ${tile.workMachines} need work, ${tile.laptops} laptops, ${tile.desktops} desktops`,
      donut: donutGradient(tile.worst, tile.count),
      center: (<><strong className="mm-name">{tile.code}</strong><span className="mm-sub">{plural(tile.count, 'machine')}</span></>),
      caption: (<><span className="mm-pill">{tile.workMachines} need work</span><span className="mm-line">{plural(tile.laptops, 'laptop')} · {plural(tile.desktops, 'desktop')}</span></>),
    })),
    {
      key: 'stash',
      shape: 'hex',
      kind: 'stash',
      at: layout.extras.stash,
      to: mapHref({ place: PLACES.STASH }),
      label: `IT Stash, ${world.stash.count} machines ready to hand out`,
      center: (<><strong className="mm-big">{world.stash.count}</strong><span className="mm-shape-name">IT Stash</span></>),
      caption: <span className="mm-line">ready to hand out</span>,
    },
    ...(world.noLocation ? [{
      key: 'nolocation',
      shape: 'diamond',
      kind: 'noLocation',
      at: layout.extras.noLocation,
      to: mapHref({ place: PLACES.NO_LOCATION }),
      label: `No location, ${world.noLocation.count} machines need a place`,
      center: <strong className="mm-big">{world.noLocation.count}</strong>,
      caption: (<><strong className="mm-caption-title">No location</strong><span className="mm-line mm-amber">give these a place</span></>),
    }] : []),
    {
      key: 'retired',
      shape: 'tomb',
      kind: 'retired',
      at: layout.extras.retired,
      to: mapHref({ place: PLACES.GRAVEYARD }),
      label: `Retired, ${world.graveyard.count} machines, not counted`,
      center: <strong className="mm-big">{world.graveyard.count}</strong>,
      caption: (<><strong className="mm-caption-title">Retired</strong><span className="mm-line">not counted</span></>),
    },
  ];

  const fleet = world.locations.reduce((sum, t) => sum + t.count, 0) + (world.noLocation?.count ?? 0);
  const laptops = world.locations.reduce((sum, t) => sum + t.laptops, 0) + (world.noLocation?.laptops ?? 0);
  const desktops = world.locations.reduce((sum, t) => sum + t.desktops, 0) + (world.noLocation?.desktops ?? 0);

  return (
    <section className="dm">
      <header className="mm-head">
        <div className="dm-head">
          <h2 className="dm-title">Where the machines are</h2>
          <p className="dm-sub">{plural(fleet, 'machine')} in use · {plural(laptops, 'laptop')} · {plural(desktops, 'desktop')}</p>
        </div>
        <GradeLegend />
      </header>
      <MindMap
        label="Places: each ring is a location, with its machines counted once at their worst part"
        hub={(<><span className="mm-hub-eyebrow">PMW fleet</span><strong className="mm-hub-count">{fleet}</strong><span className="mm-hub-sub">in use</span></>)}
        nodes={nodes}
      />
    </section>
  );
}

import { useMemo, useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import MachineCard from './MachineCard';
import PartBars from './PartBars.jsx';
import { Search } from '../../../components/ui/Icons';
import { machinesIn, searchMachines, summarise, PLACES } from '../map/zones';
import { mapHref, parentHref } from '../map/mapLinks';
import { statusOf, IN_REPAIR } from '../lifecycle/status';

const PLACE_TITLE = {
  [PLACES.STASH]: 'IT Stash', [PLACES.GRAVEYARD]: 'Graveyard', [PLACES.NO_LOCATION]: 'No location yet',
};

export default function DepartmentView({ devices, location, department, place }) {
  const navigate = useNavigate();
  const [query, setQuery] = useState('');
  const rows = useMemo(() => machinesIn(devices, { location, department, place }), [devices, location, department, place]);
  const shown = useMemo(() => searchMachines(rows, query), [rows, query]);
  const summary = useMemo(() => summarise(department ?? PLACE_TITLE[place], rows), [rows, department, place]);
  const title = place ? PLACE_TITLE[place] : department;

  const groups = place
    ? [{ key: 'all', title: null, rows: shown }]
    : [
      { key: 'use', title: 'In use', rows: shown.filter((d) => statusOf(d) !== IN_REPAIR) },
      { key: 'repair', title: 'In repair', rows: shown.filter((d) => statusOf(d) === IN_REPAIR) },
    ];

  return (
    <section className="dm" onKeyDown={(event) => {
      if (event.key === 'Escape' && event.target.tagName !== 'INPUT') navigate(parentHref({ location, department, place }));
    }}>
      <nav className="dm-crumbs" aria-label="Breadcrumb">
        <Link to={mapHref()}>Map</Link>
        {location && !place && (<><span aria-hidden="true">›</span><Link to={mapHref({ location })}>{location}</Link></>)}
        <span aria-hidden="true">›</span><span aria-current="page">{title}</span>
      </nav>
      <header className="dv-level">
        <span className="dm-eyebrow">{place ? 'Place' : `${location} · department`}</span>
        <h2 className="dv-level-title">{title}</h2>
        <div className="dv-level-stats">
          <span><strong>{summary.count}</strong> machines</span>
          <span><strong>{summary.laptops}</strong> laptops</span>
          <span><strong>{summary.desktops}</strong> desktops</span>
          {!place && <span><strong>{summary.criticalMachines}</strong> with a critical part</span>}
          {!place && <span><strong>{summary.attentionMachines}</strong> needing attention</span>}
        </div>
        {!place && <PartBars bars={summary.partBars} />}
      </header>
      <label className="dv-search">
        <Search size={18} />
        <span className="sr-only">Search {title}</span>
        <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Owner, computer name or serial" />
      </label>
      {groups.map((group) => (group.rows.length > 0 || group.key === 'use') && (
        <div key={group.key} className="dv-group">
          {group.title && <h3 className="dv-group-title">{group.title} · {group.rows.length}</h3>}
          {group.rows.length === 0
            ? <p className="dm-empty">{query ? 'Nothing matches that search.' : 'Nothing here.'}</p>
            : <div className="dv-cards">{group.rows.map((device) => <MachineCard key={device.id} device={device} />)}</div>}
        </div>
      ))}
    </section>
  );
}

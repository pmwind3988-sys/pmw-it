import { Link } from 'react-router-dom';
import { Laptop, Monitor } from '../../../components/ui/Icons';
import PartBars from './PartBars.jsx';
import { PARTS } from '../standards/defaultStandard.js';

/** One place on the map. The label says everything the tile shows, for a screen reader. */
export default function MapTile({
  to, eyebrow, title, summary, chips, variant = 'zone', cell, tileRef, onKeyDown, children,
}) {
  const criticalOn = PARTS.filter(({ key }) => summary.parts?.[key]?.Critical)
    .map(({ key, label: part }) => `${part} critical on ${summary.parts[key].Critical}`);
  const label = [
    title, `${summary.count} machines`, `${summary.laptops} laptops`, `${summary.desktops} desktops`, ...criticalOn,
  ].join(', ');
  const style = cell ? { '--col': cell.col, '--row': cell.row } : undefined;
  const body = (
    <>
      <span className="mt-head">
        <span className="mt-eyebrow">{eyebrow}</span>
        {variant === 'zone' && (summary.criticalMachines > 0
          ? <span className="mt-badge g-critical">{summary.criticalMachines} with a critical part</span>
          : summary.count > 0 && summary.attentionMachines === 0 && <span className="mt-badge g-optimal">All clear</span>)}
      </span>
      <span className="mt-title">
        <span className="mt-name">{title}</span>
        <span className="mt-count">{summary.count} machine{summary.count === 1 ? '' : 's'}</span>
      </span>
      {summary.count > 0 && <PartBars bars={summary.partBars} />}
      <span className="mt-split">
        <span><Laptop size={18} /> <strong>{summary.laptops}</strong> laptop{summary.laptops === 1 ? '' : 's'}</span>
        <span><Monitor size={18} /> <strong>{summary.desktops}</strong> desktop{summary.desktops === 1 ? '' : 's'}</span>
        {summary.other > 0 && <span><strong>{summary.other}</strong> other</span>}
      </span>
      {chips?.length > 0 && (
        <span className="mt-chips">
          {chips.slice(0, 3).map((chip) => <span key={chip.name} className="mt-chip">{chip.name} {chip.count}</span>)}
          {chips.length > 3 && <span className="mt-chip">+{chips.length - 3} more</span>}
        </span>
      )}
      {children}
    </>
  );

  if (!to) return <div className={`mt mt-${variant}`} style={style} aria-label={label}>{body}</div>;
  return (
    <Link to={to} className={`mt mt-${variant}`} style={style} aria-label={label} ref={tileRef} onKeyDown={onKeyDown}>
      {body}
    </Link>
  );
}

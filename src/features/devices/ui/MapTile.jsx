import { Link } from 'react-router-dom';
import { Laptop, Monitor } from '../../../components/ui/Icons';

const LEVEL_CLASS = {
  Critical: 'crit', 'Needs Attention': 'attn', Moderate: 'mod', Optimal: 'ok', Unknown: 'unk',
};

/** One place on the map. The label says everything the tile shows, for a screen reader. */
export default function MapTile({
  to, eyebrow, title, summary, chips, variant = 'zone', cell, tileRef, onKeyDown, children,
}) {
  const label = [
    title, `${summary.count} machines`, `${summary.laptops} laptops`, `${summary.desktops} desktops`,
    summary.critical ? `${summary.critical} critical` : null,
  ].filter(Boolean).join(', ');
  const style = cell ? { '--col': cell.col, '--row': cell.row } : undefined;
  const body = (
    <>
      <span className="mt-head">
        <span className="mt-eyebrow">{eyebrow}</span>
        {variant === 'zone' && (summary.critical > 0
          ? <span className="mt-badge mt-badge-crit">{summary.critical} critical</span>
          : summary.count > 0 && summary.attention === 0 && <span className="mt-badge mt-badge-ok">All clear</span>)}
      </span>
      <span className="mt-title">
        <span className="mt-name">{title}</span>
        <span className="mt-count">{summary.count} machine{summary.count === 1 ? '' : 's'}</span>
      </span>
      {summary.health.length > 0 && (
        <span className="mt-health" aria-hidden="true">
          {summary.health.map((part) => (
            <span key={part.level} className={`mt-seg mt-seg-${LEVEL_CLASS[part.level]}`} style={{ width: `${part.share * 100}%` }} />
          ))}
        </span>
      )}
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

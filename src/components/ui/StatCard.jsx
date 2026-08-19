import { ChevronRight } from './Icons';

/**
 * A stat card. Pass `onClick` and it becomes the button that opens the records
 * behind the number — a figure with no way to ask "which ones?" is the usual
 * complaint about a dashboard like this.
 *
 * Renders as a plain div when there is nothing to drill into, so a card never
 * advertises an affordance it does not have.
 */
export default function StatCard({ icon: Icon, label, value, unit, color, loading, onClick }) {
  const interactive = typeof onClick === 'function';

  const body = (
    <>
      <div className="stat-card-top">
        <span className="stat-card-label">{label}</span>
        {Icon && <Icon size={15} style={{ color, flexShrink: 0, marginTop: 1 }} />}
      </div>
      {loading ? (
        <div className="ui-skeleton" style={{ marginTop: 'auto' }} />
      ) : (
        <div className="stat-card-value">
          <span style={{ display: 'flex', alignItems: 'baseline', gap: 4, minWidth: 0 }}>
            <span className="stat-card-number">{value}</span>
            {unit && <span className="stat-card-unit">{unit}</span>}
          </span>
          {interactive && <ChevronRight size={15} className="stat-card-chevron" />}
        </div>
      )}
    </>
  );

  if (!interactive) return <div className="stat-card">{body}</div>;

  return (
    <button
      type="button"
      className="stat-card"
      onClick={onClick}
      aria-label={`${label}: ${loading ? 'loading' : value}${unit ? ` ${unit}` : ''}. Show the records behind this.`}
    >
      {body}
    </button>
  );
}

import { specHistory } from '../lifecycle/specHistory';

export default function SpecHistory({ changes, loading }) {
  const days = specHistory(changes);
  return (
    <section className="sh card">
      <h2 className="dd-group-title">Spec history <span className="dd-group-hint">from each scan</span></h2>
      {loading && <p className="dm-empty">Reading the history…</p>}
      {!loading && days.length === 0 && <p className="dm-empty">No changes recorded since the first scan.</p>}
      {days.map((group) => (
        <div key={group.dayLabel} className="sh-day">
          <span className="sh-date">{group.dayLabel}</span>
          <dl className="sh-rows">
            {group.rows.map((row) => (
              <div key={`${row.fieldName}-${row.newValue}`} className={row.rename ? 'sh-rename' : undefined}>
                <dt>{row.label}</dt>
                <dd><s>{row.oldValue || '—'}</s> → <strong>{row.newValue || '—'}</strong></dd>
              </div>
            ))}
          </dl>
        </div>
      ))}
    </section>
  );
}

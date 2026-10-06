import { X, AlertTriangle } from '../../../../components/ui/Icons';

/**
 * The receipt. It draws lines and knows nothing about what they are: each
 * mode hands it rows already worded, which keeps "what a scan means" in the
 * pure modules where it is tested.
 *
 * A row: { id, badge, name, sub, note, tone: 'ok'|'ask'|'bad',
 *          qty, onInc, onDec, onRemove, extra }
 */
export default function TillReceipt({ title, rows, empty }) {
  return (
    <section className="till-receipt" aria-label={title}>
      <header className="till-receipt-head">
        <span>{title}</span>
        <span>{rows.length} line{rows.length === 1 ? '' : 's'}</span>
      </header>

      {rows.length === 0 ? (
        <p className="till-receipt-empty">{empty}</p>
      ) : (
        <ol className="till-lines">
          {rows.map((row) => (
            <li key={row.id} className={`till-line till-line-${row.tone ?? 'ok'}`}>
              <div className="till-line-main">
                <span className="till-badge" aria-hidden="true">{row.badge}</span>
                <span className="till-line-text">
                  <strong>{row.name}</strong>
                  {row.sub && <span className="till-mono">{row.sub}</span>}
                  {row.note && (
                    <span className="till-note">
                      {row.tone === 'bad' && <AlertTriangle size={13} />}
                      {row.note}
                    </span>
                  )}
                </span>

                {row.onInc ? (
                  <span className="till-stepper">
                    <button type="button" onClick={row.onDec} aria-label={`One fewer ${row.name}`}>−</button>
                    <span className="till-mono">{row.qty}</span>
                    <button type="button" onClick={row.onInc} aria-label={`One more ${row.name}`}>+</button>
                  </span>
                ) : (
                  row.qty != null && <span className="till-qty till-mono">×{row.qty}</span>
                )}

                <button type="button" className="till-remove" onClick={row.onRemove} aria-label={`Remove ${row.name}`}>
                  <X size={16} />
                </button>
              </div>
              {row.extra && <div className="till-line-extra">{row.extra}</div>}
            </li>
          ))}
        </ol>
      )}
    </section>
  );
}

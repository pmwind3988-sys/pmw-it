import { Link } from 'react-router-dom';
import Button from '../../../../components/ui/Button';
import { Check, ArrowLeft } from '../../../../components/ui/Icons';

/**
 * What just happened, as a receipt, and the one button for the next job.
 * Nothing on it needs reading to carry on: "Next" is where the thumb is.
 */
export default function DoneReceipt({ done, onNext }) {
  return (
    <section className="till-done" aria-live="polite">
      <header className="till-done-head">
        <span className="till-done-tick"><Check size={22} /></span>
        <span>
          <strong>{done.title}</strong>
          <span>{done.when}</span>
        </span>
      </header>

      <div className="till-paper">
        <dl className="till-done-meta till-mono">
          {done.meta.map((entry) => (
            <div key={entry.k}><dt>{entry.k}</dt><dd>{entry.v}</dd></div>
          ))}
        </dl>
        <ul className="till-done-rows">
          {done.rows.map((row) => (
            <li key={row.id}>
              <span><strong>{row.name}</strong>{row.sub && <span className="till-mono">{row.sub}</span>}</span>
              <span className="till-mono">×{row.qty}</span>
            </li>
          ))}
        </ul>
        <div className="till-done-total till-mono"><span>TOTAL</span><span>{done.total}</span></div>
        {done.warning && <p className="till-done-warning">{done.warning}</p>}
      </div>

      <Button className="till-cta" onClick={onNext}>{done.next}</Button>
      {done.link && (
        <Link className="till-link till-done-link" to={done.link.to}>
          <ArrowLeft size={14} /> {done.link.label}
        </Link>
      )}
    </section>
  );
}

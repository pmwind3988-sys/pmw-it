import { timelineOf } from '../lifecycle/stints';
import { formatMYT } from '../../../utils/malaysiaTime';

const date = (ms) => (typeof ms === 'number' ? formatMYT(ms, 'date') : '?');

export default function OwnerHistory({ device, stints, loading }) {
  const entries = timelineOf(device, stints);
  return (
    <section className="oh card">
      <h2 className="dd-group-title">Owner history</h2>
      {loading && <p className="dm-empty">Reading the history…</p>}
      {!loading && entries.length === 0 && <p className="dm-empty">Nobody is recorded against this machine.</p>}
      <ol className="oh-list">
        {entries.map((entry) => (entry.kind === 'gap' ? (
          <li key={`gap-${entry.from}`} className="oh-item oh-gap">
            <span className="oh-dot" aria-hidden="true" />
            <span><strong>{entry.label}</strong><span className="oh-when">{date(entry.from)} – {entry.to ? date(entry.to) : 'today'}</span></span>
          </li>
        ) : (
          <li key={entry.id ?? 'legacy'} className={`oh-item${entry.current ? ' oh-current' : ''}`}>
            <span className={`oh-dot${entry.assignedOnApprox ? ' oh-dot-approx' : ''}`} aria-hidden="true" />
            <span>
              <strong>{entry.owner}</strong>
              <span className="oh-where">{[entry.location, entry.department].filter(Boolean).join(' · ')}</span>
              {entry.current && <span className="oh-now">Now</span>}
              <span className="oh-when">
                {entry.assignedOnApprox ? 'since at least ' : ''}{date(entry.assignedOn)} – {entry.endedOn ? date(entry.endedOn) : 'today'}
              </span>
              {(entry.endReason || entry.note) && (
                <span className="oh-why">{[entry.endReason, entry.note && `"${entry.note}"`].filter(Boolean).join(' · ')}</span>
              )}
            </span>
          </li>
        )))}
      </ol>
    </section>
  );
}

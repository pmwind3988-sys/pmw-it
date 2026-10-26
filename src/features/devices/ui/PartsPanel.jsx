import { useId } from 'react';
import { Card } from '../../../components/ui/Surfaces';
import { identityFieldsOf, partRowsOf } from '../partsView';

const dash = (value) => value || '—';

/**
 * The "Serial and part numbers" section: the six identity numbers, then one
 * row per memory stick, drive and screen. A serial the scan did not read says
 * so in words rather than leaving a blank.
 */
export default function PartsPanel({ device }) {
  const headingId = useId();
  const identity = identityFieldsOf(device);
  const rows = partRowsOf(device);

  return (
    <Card className="mh-section pp" aria-labelledby={headingId} role="region">
      <div className="mh-section-head">
        <h2 id={headingId} className="mh-section-title">Serial and part numbers</h2>
      </div>

      <dl className="pp-ids">
        {identity.map((field) => (
          <div key={field.key} className="pp-id">
            <dt>{field.label}</dt>
            <dd className={field.mono ? 'mh-mono' : undefined}>{dash(field.value)}</dd>
          </div>
        ))}
      </dl>

      {rows.length > 0 ? (
        <div className="pp-table" role="table" aria-label="Memory, drives and screens">
          <div className="pp-row pp-head" role="row">
            <span role="columnheader">Part</span>
            <span role="columnheader">What</span>
            <span role="columnheader">Part number</span>
            <span role="columnheader">Serial number</span>
            <span role="columnheader">Where</span>
          </div>
          {rows.map((row) => (
            <div key={row.key} className="pp-row" role="row">
              <span role="cell" className="pp-part">{row.part}</span>
              <span role="cell" className="pp-what">{row.what}</span>
              <span role="cell" className="pp-pn mh-mono">{dash(row.partNumber)}</span>
              <span role="cell" className="pp-sn mh-mono">
                {row.serial || <span className="mh-muted">not reported</span>}
              </span>
              <span role="cell" className="pp-where">{dash(row.where)}</span>
            </div>
          ))}
        </div>
      ) : (
        <p className="mh-none">No memory, drive or screen was listed by this scan.</p>
      )}

      <p className="mh-section-note">
        A part number or serial that changes between scans shows in Spec history, so a swapped stick or drive is visible.
      </p>
    </Card>
  );
}

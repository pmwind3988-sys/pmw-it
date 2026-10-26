import { useId } from 'react';
import { Card } from '../../../components/ui/Surfaces';
import { AlertTriangle } from '../../../components/ui/Icons';
import { licenceCards, licenceWarning } from '../licenceView';
import { scanDayLabel } from '../dayLabel';

/**
 * The Licences section: Windows, Office and Joined to, each a card, and the
 * amber box when a personal copy of Office sits beside the company one.
 */
export default function LicencePanel({ device }) {
  const headingId = useId();
  const cards = licenceCards(device);
  const warning = licenceWarning(device);
  const scanned = scanDayLabel(device.scannedOn);

  return (
    <Card className="mh-section lp" aria-labelledby={headingId} role="region">
      <div className="mh-section-head">
        <h2 id={headingId} className="mh-section-title">Licences</h2>
        {scanned && <span className="mh-section-note">From the scan of {scanned}</span>}
      </div>

      <ul className="lp-cards">
        {cards.map((card) => (
          <li key={card.kicker} className="lp-card">
            <div className="lp-kicker">{card.kicker}</div>
            <div className="lp-big">{card.big}</div>
            <span className={`lp-tone lp-tone-${card.toneKind}`}>{card.tone}</span>
            <div className="lp-lines">
              {card.lines.map((line, index) => (
                <p key={line} className={index === 0 ? 'lp-say' : 'lp-fact'}>{line}</p>
              ))}
            </div>
          </li>
        ))}
      </ul>

      {warning && (
        <p className="lp-warning" role="note">
          <AlertTriangle size={20} />
          <span>{warning}</span>
        </p>
      )}
    </Card>
  );
}

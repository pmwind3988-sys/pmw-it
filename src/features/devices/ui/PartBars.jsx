import { GRADE_SLUG } from '../standards/gradeColors.js';

/** Five thin bars, one per part, each split by how many machines hold each grade on it. */
export default function PartBars({ bars }) {
  return (
    <span className="pb" aria-hidden="true">
      {bars.map((bar) => (
        <span key={bar.key} className="pb-row">
          <span className="pb-label">{bar.label}</span>
          <span className="pb-track">
            {bar.segs.map((seg) => (
              <span key={seg.grade} className={`pb-seg g-${GRADE_SLUG[seg.grade]}`} style={{ width: `${seg.share * 100}%` }} />
            ))}
          </span>
        </span>
      ))}
    </span>
  );
}

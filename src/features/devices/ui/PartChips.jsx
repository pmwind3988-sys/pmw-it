import { PARTS } from '../standards/defaultStandard';
import { GRADE_SLUG } from '../standards/gradeColors';

/** One chip per part, coloured by that part's own grade. */
export default function PartChips({ device }) {
  return (
    <span className="pc">
      {PARTS.map(({ key, label }) => {
        const part = device.parts?.[key];
        const grade = part?.grade ?? 'Unknown';
        return (
          <span key={key} className={`pc-chip g-${GRADE_SLUG[grade]}`} title={part?.reason}>
            {label} · {grade}
          </span>
        );
      })}
    </span>
  );
}

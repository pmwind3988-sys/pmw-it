import { Card } from '../../../components/ui/Surfaces';
import { PARTS } from '../standards/defaultStandard';
import { GRADE_SLUG } from '../standards/gradeColors';
import { worstGradeOf, securityRiskOf } from '../health';
import { formatScalar } from '../formatValue';

// The glyph travels with the word, so a grade is never carried by colour alone.
const GLYPH = { Critical: '!', 'Needs attention': '!', Moderate: '–', Optimal: '✓', Unknown: '?' };

const gradeOf = (part) => (GLYPH[part?.grade] ? part.grade : 'Unknown');
const slugOf = (grade) => GRADE_SLUG[grade] ?? 'unknown';

/**
 * The machine's health against the standard for its desk: five part tiles,
 * the one thing to do, and the facts that sit beside the grades. Security
 * risk is shown beside it, not folded into the grades.
 */
export default function PartTable({ device }) {
  const worst = worstGradeOf(device);
  const calloutTone = worst === 'Critical' ? 'critical' : worst === 'Needs attention' ? 'attention' : 'neutral';
  const reasons = [].concat(device.riskReasons ?? []).filter(Boolean).join(' · ');
  const standardName = device.personaLabel ? ` — against the ${device.personaLabel} standard` : '';

  return (
    <Card className="mh-health">
      <div className="mh-health-head">
        <h2 className="mh-health-title">Health<span className="mh-health-sub">{standardName}</span></h2>
        <span className="mh-security" title={reasons || undefined}>
          Security risk: <strong>{securityRiskOf(device.riskLevel)}</strong>
          {typeof device.riskScore === 'number' ? ` · ${device.riskScore} / 100` : ''}
        </span>
      </div>

      {/* "Nothing to do" is the absence of a task, not one: it gets no callout. */}
      {device.actionRequired && device.actionRequired !== 'Nothing to do' && (
        <p className={`mh-callout mh-callout-${calloutTone}`}>
          <strong>What to do:</strong> {device.actionRequired}
        </p>
      )}

      <ul className="mh-parts">
        {PARTS.map(({ key, label }) => {
          const part = device.parts?.[key] ?? { grade: 'Unknown', value: '', reason: '' };
          const grade = gradeOf(part);
          const slug = slugOf(grade);
          return (
            <li key={key} className={`mh-part mh-part-${slug}`}>
              <span className="mh-badge" style={{ background: `var(--grade-${slug})`, color: `var(--grade-${slug}-ink)` }} aria-hidden="true">
                {GLYPH[grade]}
              </span>
              <span className="mh-part-name">{label}</span>
              <span className="mh-part-value">{part.value || '—'}</span>
              <span className="mh-part-grade">{grade}</span>
              <span className={`pc-chip g-${slug} mh-chip`}>
                <span aria-hidden="true">{GLYPH[grade]}</span> {grade}
              </span>
              {part.reason && <span className="mh-part-reason">{part.reason}</span>}
            </li>
          );
        })}
      </ul>

      <ul className="mh-reasons">
        {PARTS.filter(({ key }) => device.parts?.[key]?.reason).map(({ key, label }) => (
          <li key={key}><strong>{label}:</strong> {device.parts[key].reason}</li>
        ))}
      </ul>

      <dl className="mh-facts">
        <div>
          <dt>Office licence</dt>
          <dd title={formatScalar(device.licenseNote) === '—' ? undefined : formatScalar(device.licenseNote)}>
            {formatScalar(device.licenseStatus)}
          </dd>
        </div>
        <div>
          <dt>Antivirus</dt>
          <dd>{formatScalar(device.antivirusStatus)}</dd>
        </div>
        <div>
          <dt>Server link</dt>
          <dd title={formatScalar(device.networkNote) === '—' ? undefined : formatScalar(device.networkNote)}>
            {device.serverDependent ? formatScalar(device.networkRisk) : 'Not server-bound'}
          </dd>
        </div>
        <div>
          <dt>Form factor</dt>
          <dd>
            {formatScalar(device.suggestedFormFactor)}
            {device.formFactorNote && <span className="mh-note"> {formatScalar(device.formFactorNote)}</span>}
          </dd>
        </div>
      </dl>
    </Card>
  );
}

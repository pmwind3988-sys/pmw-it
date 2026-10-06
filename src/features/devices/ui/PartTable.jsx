import { Card } from '../../../components/ui/Surfaces';
import { PARTS } from '../standards/defaultStandard';
import { GRADE_SLUG } from '../standards/gradeColors';

/** Every part against the standard for this machine's desk, with the reason for each grade. */
export default function PartTable({ device }) {
  return (
    <Card className="dd-fit">
      <h2 className="dd-group-title">
        Parts against the {device.personaLabel} standard
        <span className="dd-group-hint">{device.personaBlurb}</span>
      </h2>
      <div className="pt-scroll">
        <table className="pt">
          <thead>
            <tr><th scope="col">Part</th><th scope="col">This machine</th><th scope="col">Grade</th><th scope="col">Why</th></tr>
          </thead>
          <tbody>
            {PARTS.map(({ key, label }) => {
              const part = device.parts?.[key] ?? { grade: 'Unknown', value: '—', reason: '' };
              return (
                <tr key={key}>
                  <th scope="row">{label}</th>
                  <td>{part.value || '—'}</td>
                  <td><span className={`pc-chip g-${GRADE_SLUG[part.grade]}`}>{part.grade}</span></td>
                  <td>{part.reason}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <dl className="dd-fit-facts">
        <div><dt>Action</dt><dd>{device.actionRequired ?? '—'}</dd></div>
        <div>
          <dt>Suggested form factor</dt>
          <dd>{device.suggestedFormFactor ?? '—'}<span className="dd-fit-note">{device.formFactorNote}</span></dd>
        </div>
        <div>
          <dt>Office licence</dt>
          <dd>{device.licenseStatus ?? '—'}<span className="dd-fit-note">{device.licenseNote}</span></dd>
        </div>
        <div>
          <dt>Server link</dt>
          <dd>
            {device.serverDependent ? device.networkRisk : 'Not server-bound'}
            <span className="dd-fit-note">{device.networkNote}</span>
          </dd>
        </div>
      </dl>
    </Card>
  );
}

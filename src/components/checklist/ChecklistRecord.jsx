import Logo from '../Logo';
import { FORM_MODES, fieldsFor } from '../../features/forms/checklistForm';
import { FIELD_LABELS, describeValue } from '../../features/forms/describe';
import { formatMYT } from '../../utils/malaysiaTime';

/**
 * A signed checklist as a document: what the employee keeps, and what prints
 * to one A4 page. Nothing on it is an input — once signed, the record is the
 * record, and a page that still looked editable would suggest otherwise.
 *
 * The print rules live in `public.css` under `@media print`.
 */

const DETAILS = ['employeeName', 'employeeNo', 'position', 'entity', 'formDate'];

function Value({ field, value }) {
  const shown = describeValue(field, value);
  if (Array.isArray(shown)) {
    return shown.length
      ? <ul className="cr-list">{shown.map((line) => <li key={line}>{line}</li>)}</ul>
      : <span className="cr-empty">None</span>;
  }
  return shown ? <span className="cr-text">{shown}</span> : <span className="cr-empty">—</span>;
}

export default function ChecklistRecord({ formMode, values = {}, signature, signedOn }) {
  const mode = FORM_MODES.find((entry) => entry.value === formMode);
  const shown = new Set(fieldsFor(formMode));
  const listField = shown.has('items') ? 'items' : 'checkedItems';
  const signedAt = Date.parse(signedOn);

  return (
    <article className="cr-doc" aria-label="Signed asset checklist">
      <header className="cr-head">
        <Logo size={44} />
        <div>
          <p className="cr-kicker">PMW Group · IT</p>
          <h1 className="cr-title">IT Asset Tracking Form</h1>
        </div>
        <div className="cr-mode">
          <span className="cr-mode-label">{mode?.label ?? formMode}</span>
          {mode?.description && <span className="cr-mode-desc">{mode.description}</span>}
        </div>
      </header>

      <dl className="cr-grid">
        {DETAILS.map((field) => (
          <div key={field} className="cr-cell">
            <dt>{FIELD_LABELS[field]}</dt>
            <dd><Value field={field} value={values[field]} /></dd>
          </div>
        ))}
      </dl>

      <section className="cr-section">
        <h2>{FIELD_LABELS[listField]}</h2>
        <Value field={listField} value={values[listField]} />
      </section>

      <div className="cr-pair">
        <section className="cr-section">
          <h2>{FIELD_LABELS.serialNumbers}</h2>
          <Value field="serialNumbers" value={values.serialNumbers} />
        </section>
        <section className="cr-section">
          <h2>{FIELD_LABELS.otherRemarks}</h2>
          <Value field="otherRemarks" value={values.otherRemarks} />
        </section>
      </div>

      <footer className="cr-sign">
        <div className="cr-sign-box">
          {signature
            ? <img src={signature} alt="Employee's signature" />
            : <span className="cr-empty">Signature on file with IT</span>}
        </div>
        <div className="cr-sign-meta">
          <span className="cr-sign-name">{values.employeeName}</span>
          {Number.isFinite(signedAt) && (
            <span>Signed {formatMYT(signedAt, 'datetime12')} (Malaysia time)</span>
          )}
        </div>
      </footer>
    </article>
  );
}

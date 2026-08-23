/**
 * A labelled form field.
 *
 * The error sits under the control and is announced, rather than being a colour
 * change alone — a red border tells somebody that something is wrong but not
 * what, and tells a screen reader nothing at all.
 */
export default function Field({
  label, htmlFor, required, help, error, children, wide = false,
}) {
  return (
    <div className={`ff-field${wide ? ' ff-field-wide' : ''}${error ? ' ff-field-bad' : ''}`}>
      {label && (
        <label className="ff-label" htmlFor={htmlFor}>
          {label}
          {required && <span className="ff-req" aria-hidden="true">*</span>}
          {required && <span className="ff-sr">(required)</span>}
        </label>
      )}
      {help && <p className="ff-help">{help}</p>}
      {children}
      {error && (
        <p className="ff-error" role="alert">{error}</p>
      )}
    </div>
  );
}

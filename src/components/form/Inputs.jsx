/**
 * The plain controlled inputs the two forms are built from.
 *
 * Every one of them takes `value` and `onChange(value)` — the raw value, not
 * the event. The forms hold a plain object of values, and a control that hands
 * back an event makes every call site unwrap it identically.
 */

const invalid = (error) => (error ? { 'aria-invalid': 'true' } : null);

export function TextInput({ id, value, onChange, error, ...rest }) {
  return (
    <input
      id={id}
      className="ff-input"
      value={value ?? ''}
      onChange={(event) => onChange(event.target.value)}
      {...invalid(error)}
      {...rest}
    />
  );
}

export function TextArea({ id, value, onChange, error, rows = 3, ...rest }) {
  return (
    <textarea
      id={id}
      className="ff-input ff-textarea"
      rows={rows}
      value={value ?? ''}
      onChange={(event) => onChange(event.target.value)}
      {...invalid(error)}
      {...rest}
    />
  );
}

/**
 * Kept as a STRING rather than coerced on every keystroke: parsing as the user
 * types turns a half-deleted "1" into 0 or NaN and fights them. The form
 * coerces once, on submit.
 */
export function NumberInput({ id, value, onChange, error, min = 1, ...rest }) {
  return (
    <input
      id={id}
      type="number"
      inputMode="numeric"
      className="ff-input ff-number"
      min={min}
      value={value ?? ''}
      onChange={(event) => onChange(event.target.value)}
      {...invalid(error)}
      {...rest}
    />
  );
}

export function DateInput({ id, value, onChange, error, ...rest }) {
  return (
    <input
      id={id}
      type="date"
      className="ff-input"
      value={value ?? ''}
      onChange={(event) => onChange(event.target.value)}
      {...invalid(error)}
      {...rest}
    />
  );
}

/**
 * `options` is a list of strings, or of `{ value, label }`.
 *
 * The empty option is always offered so that "nothing chosen" is a state the
 * user can see and return to — a select that silently starts on its first
 * option collects that option from everybody who never touched it.
 */
export function SelectInput({ id, value, onChange, options = [], error, placeholder = 'Choose…', ...rest }) {
  return (
    <select
      id={id}
      className="ff-input ff-select"
      value={value ?? ''}
      onChange={(event) => onChange(event.target.value)}
      {...invalid(error)}
      {...rest}
    >
      <option value="">{placeholder}</option>
      {options.map((option) => {
        const entry = typeof option === 'string' ? { value: option, label: option } : option;
        return <option key={entry.value} value={entry.value}>{entry.label}</option>;
      })}
    </select>
  );
}

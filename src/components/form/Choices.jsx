import { Check } from '../ui/Icons';

/**
 * The two multi-option controls: big radio cards for a choice that decides the
 * shape of the rest of the form, and a tick list for "which of these".
 */

/**
 * Real radio inputs underneath, visually hidden.
 *
 * A row of `<button>`s would look identical and lose every keyboard behaviour a
 * radio group has for free — arrow keys moving between options, and a screen
 * reader saying "2 of 3".
 */
export function RadioCards({ name, value, onChange, options, error }) {
  return (
    <div className="ff-cards" role="radiogroup" aria-invalid={error ? 'true' : undefined}>
      {options.map((option) => (
        <label
          key={option.value}
          className={`ff-card${value === option.value ? ' ff-card-on' : ''}`}
        >
          <input
            type="radio"
            name={name}
            className="ff-sr-input"
            value={option.value}
            checked={value === option.value}
            onChange={() => onChange(option.value)}
          />
          <span className="ff-card-label">{option.label}</span>
          {option.description && (
            <span className="ff-card-desc">{option.description}</span>
          )}
        </label>
      ))}
    </div>
  );
}

/** `value` is the list of ticked items; `onChange` gets the new list. */
export function CheckList({ value = [], onChange, options, columns = true }) {
  const toggle = (option) => {
    onChange(value.includes(option)
      ? value.filter((entry) => entry !== option)
      : [...value, option]);
  };

  return (
    <div className={`ff-checks${columns ? ' ff-checks-grid' : ''}`}>
      {options.map((option) => {
        const on = value.includes(option);
        return (
          <label key={option} className={`ff-check${on ? ' ff-check-on' : ''}`}>
            <input
              type="checkbox"
              className="ff-sr-input"
              checked={on}
              onChange={() => toggle(option)}
            />
            <span className="ff-check-box" aria-hidden="true">
              {on && <Check size={12} />}
            </span>
            <span>{option}</span>
          </label>
        );
      })}
    </div>
  );
}

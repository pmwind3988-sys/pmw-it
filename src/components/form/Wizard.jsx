import Button from '../ui/Button';
import { ArrowLeft, Check } from '../ui/Icons';

/**
 * The steps across the top, and the buttons at the bottom.
 *
 * `onNext` is expected to return false when the step does not validate, so the
 * wizard stays where it is. The validation itself belongs to the form, not to
 * this: what makes a step complete is a question about the form's own fields.
 */
export default function Wizard({
  steps, current, onBack, onNext, onSubmit, submitting, submitLabel = 'Submit', children,
}) {
  const last = current === steps.length - 1;

  return (
    <div className="ff-wizard">
      <ol className="ff-steps">
        {steps.map((step, index) => (
          <li
            key={step}
            className={`ff-step${index === current ? ' ff-step-on' : ''}${index < current ? ' ff-step-done' : ''}`}
            aria-current={index === current ? 'step' : undefined}
          >
            <span className="ff-step-dot">
              {index < current ? <Check size={12} /> : index + 1}
            </span>
            <span className="ff-step-name">{step}</span>
          </li>
        ))}
      </ol>

      <div className="ff-wizard-body">{children}</div>

      <div className="ff-wizard-foot">
        {current > 0 && (
          <Button variant="ghost" icon={ArrowLeft} onClick={onBack} disabled={submitting}>
            Back
          </Button>
        )}
        {last ? (
          <Button icon={Check} onClick={onSubmit} disabled={submitting}>
            {submitting ? 'Submitting…' : submitLabel}
          </Button>
        ) : (
          <Button onClick={onNext}>Next</Button>
        )}
      </div>
    </div>
  );
}

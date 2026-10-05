import { ANSWERS } from '../lifecycle/replacements';
import { formatMYT } from '../../../utils/malaysiaTime';

const CHOICES = [
  [ANSWERS.STASH, 'Old one to IT Stash', (name) => `${name} becomes a spare, ready for someone else`],
  [ANSWERS.RETIRE, 'Retire old one', () => 'Goes to the Graveyard. Its record is kept'],
  [ANSWERS.KEEP, 'Keep both', (name, owner) => `${owner} really uses two machines`],
];

/** One "is this a replacement?" question. Nothing is preselected: the answer is the point. */
export default function ReplacementPrompt({ prompt, answer, onAnswer }) {
  const since = prompt.old.since ? ` since at least ${formatMYT(prompt.old.since, 'date')}` : '';
  return (
    <fieldset className="rp">
      <legend className="rp-legend">
        <strong>
          {prompt.owner} already has <span className="rp-code">{prompt.old.computerName}</span>
          {' '}({prompt.old.deviceType ?? 'machine'}, {String(prompt.old.status).toLowerCase()}{since}).
        </strong>
        <span>Is <span className="rp-code">{prompt.incomingName}</span> replacing it?</span>
      </legend>
      <div className="rp-choices">
        {CHOICES.map(([value, label, hint]) => (
          <label key={value} className={`rp-choice${answer === value ? ' rp-choice-on' : ''}`}>
            <input
              type="radio"
              name={`rp-${prompt.key}`}
              value={value}
              checked={answer === value}
              onChange={() => onAnswer(prompt.key, value)}
            />
            <span>
              <strong>{label}</strong>
              <span className="rp-hint">{hint(prompt.old.computerName, prompt.owner)}</span>
            </span>
          </label>
        ))}
      </div>
    </fieldset>
  );
}

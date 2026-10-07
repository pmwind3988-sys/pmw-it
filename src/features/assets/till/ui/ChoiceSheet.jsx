import TillSheet from './TillSheet';

/**
 * A short list of big choices: how a handover is signed, what a return
 * records. `options`: [{ key, title, body, onPick }]. The first is the usual
 * answer and sits where the thumb is.
 */
export default function ChoiceSheet({ title, lede, options, onClose }) {
  return (
    <TillSheet title={title} onClose={onClose}>
      {lede && <p className="till-sheet-lede">{lede}</p>}
      <div className="till-options">
        {options.map((option, index) => (
          <button key={option.key} type="button" className="till-option" onClick={option.onPick}>
            <span className="till-option-n till-mono">{index + 1}</span>
            <span className="till-option-text">
              <strong>{option.title}</strong>
              <span>{option.body}</span>
            </span>
          </button>
        ))}
      </div>
    </TillSheet>
  );
}

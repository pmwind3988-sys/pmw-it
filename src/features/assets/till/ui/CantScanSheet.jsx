import TillSheet from './TillSheet';

/**
 * The ways round a barcode that will not read, in the order worth trying.
 * "Nothing on it" is offered only when stocking in: anything being handed out
 * or taken back is already in the register, so it can always be typed.
 */
export default function CantScanSheet({ mode, onClose, onReadLabel, onType, onNoCode }) {
  const options = [
    {
      key: 'label',
      title: 'Read the printed label',
      body: 'The camera reads the words on the sticker: model, serial number, part number.',
      onPick: onReadLabel,
    },
    {
      key: 'type',
      title: 'Type it to find it',
      body: mode === 'back'
        ? 'A few letters of the serial, the PMW label, the model — or who had it.'
        : 'A few letters of the serial, the PMW label, or the model. Matches show as you type.',
      onPick: onType,
    },
  ];
  if (mode === 'in') {
    options.push({
      key: 'none',
      title: 'It has nothing on it',
      body: 'Say what it is. A laptop, monitor or dock gets a new PMW label to stick on.',
      onPick: onNoCode,
    });
  }

  return (
    <TillSheet title="Barcode won’t read?" onClose={onClose}>
      <p className="till-sheet-lede">Try these in order. Each one is quicker than the next.</p>
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

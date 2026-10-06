import { useState } from 'react';
import TillSheet from './TillSheet';
import Button from '../../../../components/ui/Button';
import { trackingModeFor, TRACKED } from '../../assetKinds';
import { Plus } from '../../../../components/ui/Icons';

/**
 * Something with nothing on it to scan. What it is decides whether it is one
 * item (and gets a label so it can be found again) or a count.
 */
export default function NoCodeSheet({ categories, tag, onCancel, onAdd }) {
  const [category, setCategory] = useState('');
  const [model, setModel] = useState('');
  const [quantity, setQuantity] = useState(1);

  const tracked = category && trackingModeFor(category) === TRACKED;

  return (
    <TillSheet title="Nothing on it to scan" onClose={onCancel}>
      <p className="till-sheet-lede">Say what it is. It goes on the receipt like anything else.</p>

      <fieldset className="till-fieldset">
        <legend>What is it?</legend>
        <div className="till-chips">
          {categories.map((name) => (
            <button
              key={name}
              type="button"
              className={category === name ? 'till-chip on' : 'till-chip'}
              aria-pressed={category === name}
              onClick={() => setCategory(name)}
            >
              {name}
            </button>
          ))}
        </div>
      </fieldset>

      <label className="till-field">
        <span>Make and model, if you know it</span>
        <input value={model} onChange={(event) => setModel(event.target.value)} placeholder="e.g. Ugreen USB-C hub" />
      </label>

      {category && !tracked && (
        <label className="till-field till-field-row">
          <span>How many?</span>
          <input
            type="number"
            inputMode="numeric"
            min="1"
            value={quantity}
            onChange={(event) => setQuantity(event.target.value)}
          />
        </label>
      )}

      {tracked && (
        <div className="till-newtag">
          <span className="till-mono">{tag}</span>
          <span>Its new label. Write it on a sticker and put it on — scanning or typing it finds this from then on.</span>
        </div>
      )}

      <Button
        icon={Plus}
        className="till-cta"
        disabled={!category}
        onClick={() => onAdd({ category, model, quantity, tag })}
      >
        {category ? 'Add to the receipt' : 'Pick what it is first'}
      </Button>
    </TillSheet>
  );
}

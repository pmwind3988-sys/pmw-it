import { useMemo, useRef, useState } from 'react';
import Button from '../../../../components/ui/Button';
import { Plus, ScanLine, Camera } from '../../../../components/ui/Icons';
import { trackingModeFor, TRACKED } from '../../assetKinds';
import PhotoInput from '../../ui/PhotoInput';
import CodeScanSheet from '../../ui/CodeScanSheet';
import TextScanSheet from '../../ui/TextScanSheet';
import { recentModels, knownLocations, COUNT_RESULT } from '../count';

/**
 * Counting what is already here. Built for walking a room: where you are is
 * set once, then each thing is a category tap, a model tap and a number — or,
 * for a thing tracked one by one, its serial (typed, scanned, or read off its
 * sticker) or an honest "no serial" with a photo.
 *
 * After an add the category, model and place STAY, because the next thing on
 * the shelf is usually another of the same.
 */
const CONDITIONS = ['Good', 'Fair', 'Faulty'];

export default function CountPanel({ assets, drafts, categories, onAdd }) {
  const [location, setLocation] = useState('');
  const [category, setCategory] = useState('');
  const [manufacturer, setManufacturer] = useState('');
  const [model, setModel] = useState('');
  const [quantity, setQuantity] = useState(1);
  const [serial, setSerial] = useState('');
  const [noSerial, setNoSerial] = useState(false);
  const [condition, setCondition] = useState('Good');
  const [photoId, setPhotoId] = useState(null);
  const [sheet, setSheet] = useState(null);
  const modelStepRef = useRef(null);

  const tracked = Boolean(category) && trackingModeFor(category) === TRACKED;
  const places = useMemo(() => knownLocations(assets, drafts), [assets, drafts]);
  const models = useMemo(
    () => (category ? recentModels(category, assets, drafts) : []),
    [category, assets, drafts],
  );

  const pickCategory = (name) => {
    // The next step is below a long grid on a phone; bring it into view.
    setTimeout(() => modelStepRef.current?.scrollIntoView({ block: 'start', behavior: 'smooth' }), 0);
    if (name === category) return;
    setCategory(name);
    setManufacturer('');
    setModel('');
    setQuantity(1);
    setSerial('');
    setNoSerial(false);
    setPhotoId(null);
  };

  const ready = Boolean(category && model.trim() && (!tracked || serial.trim() || noSerial));

  const add = () => {
    const result = onAdd({
      category, manufacturer, model, quantity, location, condition,
      serialNumber: tracked ? serial : '', noSerial: tracked && noSerial, photoId: tracked ? photoId : null,
    });
    if (result === COUNT_RESULT.ADDED || result === COUNT_RESULT.COUNTED) {
      setQuantity(1);
      setSerial('');
      setNoSerial(false);
      setPhotoId(null);
    }
  };

  const label = !category ? 'Pick what it is'
    : !model.trim() ? 'Pick or type the model'
      : tracked && !serial.trim() && !noSerial ? 'Add its serial, or tick “no serial”'
        : tracked ? `Add ${[manufacturer, model].filter(Boolean).join(' ')}`
          : `Add ${quantity} × ${[manufacturer, model].filter(Boolean).join(' ')}`;

  return (
    <section className="till-count-panel" aria-label="Catch-up count">
      <div className="till-count-step">
        <label className="till-field" htmlFor="count-where">
          <span className="till-kicker">1 · Where are you counting?</span>
        </label>
        <input
          id="count-where"
          className="till-count-input"
          value={location}
          onChange={(event) => setLocation(event.target.value)}
          placeholder="e.g. F1 IT store, HR office"
        />
        {places.length > 0 && (
          <div className="till-chips">
            {places.map((place) => (
              <button key={place} type="button" className={place === location ? 'till-chip on' : 'till-chip'} aria-pressed={place === location} onClick={() => setLocation(place)}>{place}</button>
            ))}
          </div>
        )}
      </div>

      <div className="till-count-step">
        <span className="till-kicker">2 · What is it?</span>
        <div className="till-cats">
          {categories.map((name) => (
            <button
              key={name}
              type="button"
              className={name === category ? 'till-cat on' : 'till-cat'}
              aria-pressed={name === category}
              onClick={() => pickCategory(name)}
            >
              <strong>{name}</strong>
              <span>{trackingModeFor(name) === TRACKED ? 'one by one' : 'counted'}</span>
            </button>
          ))}
        </div>
      </div>

      {category && (
        <div className="till-count-step" ref={modelStepRef}>
          <span className="till-kicker">3 · Which model?</span>
          {models.length > 0 && (
            <div className="till-chips">
              {models.map((entry) => {
                const on = entry.model.toLowerCase() === model.trim().toLowerCase()
                  && entry.manufacturer.toLowerCase() === manufacturer.trim().toLowerCase();
                return (
                  <button
                    key={`${entry.manufacturer}|${entry.model}`}
                    type="button"
                    className={on ? 'till-chip on' : 'till-chip'}
                    aria-pressed={on}
                    onClick={() => { setManufacturer(entry.manufacturer); setModel(entry.model); }}
                  >
                    {[entry.manufacturer, entry.model].filter(Boolean).join(' ')}
                  </button>
                );
              })}
            </div>
          )}
          <div className="till-count-pair">
            <label className="till-field"><span>Make</span>
              <input value={manufacturer} onChange={(event) => setManufacturer(event.target.value)} placeholder="Dell" />
            </label>
            <label className="till-field"><span>Model</span>
              <input value={model} onChange={(event) => setModel(event.target.value)} placeholder="P2422H" />
            </label>
          </div>
        </div>
      )}

      {category && !tracked && (
        <div className="till-count-step">
          <span className="till-kicker">4 · How many here?</span>
          <div className="till-bigstep">
            <button type="button" onClick={() => setQuantity((n) => Math.max(1, n - 1))} aria-label="One fewer">−</button>
            <input
              type="number"
              inputMode="numeric"
              min="1"
              aria-label="How many"
              className="till-mono"
              value={quantity}
              onChange={(event) => setQuantity(Math.max(1, Math.floor(Number(event.target.value) || 1)))}
            />
            <button type="button" onClick={() => setQuantity((n) => n + 1)} aria-label="One more">+</button>
            <button type="button" className="till-chip" onClick={() => setQuantity((n) => n + 5)}>+5</button>
            <button type="button" className="till-chip" onClick={() => setQuantity((n) => n + 10)}>+10</button>
          </div>
        </div>
      )}

      {tracked && (
        <div className="till-count-step">
          <span className="till-kicker">4 · Its serial number</span>
          <div className="till-entry till-entry-light">
            <input
              value={serial}
              onChange={(event) => { setSerial(event.target.value); if (event.target.value) setNoSerial(false); }}
              placeholder="Type it, or use a button"
              aria-label="Serial number"
              autoCapitalize="characters"
              spellCheck={false}
              disabled={noSerial}
            />
          </div>
          <div className="till-chips">
            <button type="button" className="till-chip" onClick={() => setSheet('code')} disabled={noSerial}><ScanLine size={14} /> Scan its barcode</button>
            <button type="button" className="till-chip" onClick={() => setSheet('text')} disabled={noSerial}><Camera size={14} /> Read it off the sticker or screen</button>
          </div>
          <label className="till-check">
            <input type="checkbox" checked={noSerial} onChange={(event) => { setNoSerial(event.target.checked); if (event.target.checked) setSerial(''); }} />
            No serial anywhere on it — record it with a photo instead
          </label>

          <div className="till-chips" role="group" aria-label="Condition">
            {CONDITIONS.map((name) => (
              <button key={name} type="button" className={condition === name ? 'till-chip on' : 'till-chip'} aria-pressed={condition === name} onClick={() => setCondition(name)}>{name}</button>
            ))}
          </div>

          <PhotoInput photoId={photoId} onChange={setPhotoId} label={noSerial ? 'Photo (recommended — it is how this one will be told apart)' : 'Photo (optional)'} compact />
        </div>
      )}

      <Button icon={Plus} className="till-cta" disabled={!ready} onClick={add}>{label}</Button>

      {sheet === 'code' && (
        <CodeScanSheet
          title="Scan its serial"
          onCancel={() => setSheet(null)}
          onUse={(values) => {
            const found = values.serialNumber || values.assetTag || '';
            if (found) { setSerial(found); setNoSerial(false); }
            setSheet(null);
          }}
        />
      )}
      {sheet === 'text' && (
        <TextScanSheet
          title="Read the serial off it"
          onCancel={() => setSheet(null)}
          onUse={(values) => {
            if (values.manufacturer && !manufacturer) setManufacturer(values.manufacturer);
            if (values.model && !model) setModel(values.model);
            if (values.serialNumber) {
              setSerial(values.serialNumber);
              setNoSerial(false);
              setSheet(null);
            }
          }}
        />
      )}
    </section>
  );
}

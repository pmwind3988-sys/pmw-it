import { useEffect, useState } from 'react';
import {
  ChevronLeft, ChevronRight, Boxes, Check, Camera,
} from '../../../components/ui/Icons';
import {
  UNIT_FIELDS, PER_UNIT_CODES, setUnitField, isBlankUnit, filledCount,
} from '../units';
import { labelFor } from '../scan/fieldLabels';
import ScanField from './ScanField';
import CodeScanSheet from './CodeScanSheet';
import AssetPhoto from './AssetPhoto';
import { prefetchSharePointImage } from './useSharePointImage';
import { useSharePointToken } from '../../../hooks/useRequests';
import PhotoInput from './PhotoInput';

/**
 * A bulk row, one physical item at a time.
 *
 * The register says "2 tabs" because that is what was bought. This is where
 * the second tab gets its own serial number, its own sticker label and its own
 * "screen cracked" — without splitting the line into two rows and losing the
 * fact that it was one delivery of two identical things.
 *
 * It pages rather than scrolls. Five identical cards stacked down a phone
 * screen are impossible to tell apart, and the question being answered is
 * always about ONE of them.
 *
 * The arrows are the only way between them. It used to turn on a sideways
 * swipe as well, which sat on top of a card full of text boxes: selecting a
 * serial number to correct it, or dragging the pager itself, turned the page
 * and took the field being worked on with it.
 */

export default function UnitPager({
  units, onChange, siteUrl, rowPhoto, poPhoto,
}) {
  const [at, setAt] = useState(0);
  const [scanning, setScanning] = useState(false);
  // What the camera read but would not write over something already typed.
  // Offered rather than dropped: the code on the box is evidence, and the
  // value sitting in the field might be last week's typo.
  const [heldBack, setHeldBack] = useState([]);

  const count = units.length;
  // Clamped during render rather than in an effect: lowering the quantity from
  // five to two while sitting on unit five must not paint a frame of nothing.
  const index = Math.min(at, count - 1);
  const unit = units[index];

  const go = (to) => {
    setAt(Math.max(0, Math.min(count - 1, to)));
    // Codes read off the box in your hand say nothing about the next box.
    setHeldBack([]);
  };

  // The neighbouring items' photographs are fetched while this one is being
  // read. Paging through a delivery of ten is the one place where the wait for
  // a picture is felt every few seconds, and by the time the arrow is pressed
  // the next one is already here.
  const getToken = useSharePointToken();
  useEffect(() => {
    for (const step of [1, -1]) {
      const near = units[index + step];
      if (near?.photoUrl) prefetchSharePointImage(siteUrl, near.photoUrl, getToken);
    }
  }, [units, index, siteUrl, getToken]);

  const put = (field, value) => onChange(setUnitField(units, unit.index, field, value));
  const set = (field) => (event) => put(field, event.target.value);

  /**
   * What the barcodes said, written onto THIS item.
   *
   * Only into fields that are empty. A scan is evidence about the box in
   * frame; a value already in the field was put there by somebody holding the
   * thing, and silently overwriting that is how a corrected serial number goes
   * back to being wrong.
   */
  const useScan = (reading) => {
    let next = units;
    const held = [];

    for (const field of PER_UNIT_CODES) {
      const found = reading[field];
      if (!found) continue;

      const current = String(unit[field] ?? '').trim();
      if (!current) next = setUnitField(next, unit.index, field, found);
      else if (current !== found) held.push({ field, value: found });
    }

    onChange(next);
    setHeldBack(held);
    setScanning(false);
  };

  const recorded = filledCount(units);

  return (
    <div className="as-units">
      <div className="as-units-bar">
        <button
          type="button"
          className="as-iconbtn"
          onClick={() => go(index - 1)}
          disabled={index === 0}
          aria-label="Previous item"
        >
          <ChevronLeft size={15} />
        </button>

        <div className="as-units-which">
          <strong>Item {index + 1} of {count}</strong>
          <span className="as-sub">
            {recorded === 0
              ? 'Nothing recorded against any of them yet'
              : `${recorded} of ${count} filled in`}
          </span>
        </div>

        <button
          type="button"
          className="as-iconbtn"
          onClick={() => go(index + 1)}
          disabled={index === count - 1}
          aria-label="Next item"
        >
          <ChevronRight size={15} />
        </button>
      </div>

      {/* Dots, so twenty cables do not need twenty presses to reach the one
          that matters, and so a filled unit is findable at a glance. */}
      {count > 1 && (
        <div className="as-units-dots" role="tablist" aria-label="Items on this row">
          {units.map((entry, position) => (
            <button
              key={entry.index}
              type="button"
              role="tab"
              aria-selected={position === index}
              aria-label={`Item ${position + 1}`}
              className={[
                'as-units-dot',
                position === index ? 'is-at' : '',
                isBlankUnit(entry) ? '' : 'is-filled',
              ].filter(Boolean).join(' ')}
              onClick={() => go(position)}
            />
          ))}
        </div>
      )}

      {/* Every picture that bears on THIS item, together: the one taken of it,
          the one taken of the delivery it came in, and the delivery order it
          was listed on. They were always three separate things kept in three
          separate places, which meant nobody looked at any of them. */}
      <div className="as-unit-shots">
        <figure className="as-shot">
          <PhotoInput
            photoId={unit.photoId || null}
            onChange={(photoId) => put('photoId', photoId ?? '')}
            label={`Photo of item ${index + 1}`}
            compact
          />
          <figcaption className="as-shot-caption">
            <Camera size={11} /> This one
          </figcaption>
        </figure>

        {/* The one already saved against this item. Shown beside the camera
            button rather than instead of it, so retaking is always one press
            and the picture being replaced is visible while you do it. */}
        {unit.photoUrl && !unit.photoId && (
          <AssetPhoto
            siteUrl={siteUrl}
            stored={unit.photoUrl}
            alt={`Item ${index + 1}`}
            caption={`Item ${index + 1}, as saved`}
            thumb
          />
        )}

        {rowPhoto && (
          <AssetPhoto
            siteUrl={siteUrl}
            stored={rowPhoto}
            alt="The delivery"
            caption="The whole line"
            thumb
          />
        )}

        {poPhoto && (
          <AssetPhoto
            siteUrl={siteUrl}
            stored={poPhoto}
            alt="Delivery order"
            caption="DO / PO"
            thumb
          />
        )}
      </div>

      <div className="as-form">
        {UNIT_FIELDS.map((field) => {
          const control = field.options ? (
            <select value={unit[field.key]} onChange={set(field.key)}>
              {/* Empty is "nobody has said", not a value. The row has no
                  condition of its own to fall back to — it is a count of
                  things, and only a thing can be faulty. */}
              <option value="">— not recorded</option>
              {field.options.map((option) => (
                <option key={option} value={option}>{option}</option>
              ))}
            </select>
          ) : field.multiline ? (
            <textarea rows={2} value={unit[field.key]} onChange={set(field.key)} />
          ) : (
            <input
              value={unit[field.key]}
              onChange={set(field.key)}
              placeholder={field.key === 'serialNumber' ? 'The serial on this one' : ''}
            />
          );

          return (
            <label className="as-field" key={field.key}>
              <span className="as-field-label">{field.label}</span>
              {/* The camera sits on the fields a barcode can honestly fill.
                  Nothing on a sticker says where the thing is kept or what
                  somebody thinks of the state it arrived in. */}
              {PER_UNIT_CODES.includes(field.key) ? (
                <ScanField label={field.label} onScan={() => setScanning(true)}>
                  {control}
                </ScanField>
              ) : control}
            </label>
          );
        })}
      </div>

      {heldBack.length > 0 && (
        <ul className="as-heldback">
          {heldBack.map((entry) => (
            <li key={entry.field}>
              <span className="as-heldback-label">{labelFor(entry.field)}</span>
              <span className="as-heldback-value">{entry.value}</span>
              <button
                type="button"
                className="as-heldback-take"
                onClick={() => {
                  put(entry.field, entry.value);
                  setHeldBack(heldBack.filter((held) => held.field !== entry.field));
                }}
              >
                Use this instead
              </button>
            </li>
          ))}
        </ul>
      )}

      {scanning && (
        <CodeScanSheet
          title={`Scan item ${index + 1} of ${count}`}
          onCancel={() => setScanning(false)}
          onUse={useScan}
        />
      )}

      <p className="as-units-foot">
        <Boxes size={13} />
        <span>
          Use the arrows to reach the other {count === 2 ? 'one' : 'ones'}.
          {' '}Changes to every item are saved together with Save changes.
        </span>
        {!isBlankUnit(unit) && <Check size={13} className="as-units-tick" />}
      </p>
    </div>
  );
}

import { CATEGORIES, CONDITIONS, TRACKED, BULK } from '../assetKinds';
import { setDraftField, swapSerialAndPart } from '../draft/draftAsset';
import { Trash2, Barcode, AlertTriangle, RefreshCw } from '../../../components/ui/Icons';
import PhotoInput from './PhotoInput';

/**
 * One row of a delivery, open for correction.
 *
 * A card rather than a table row on purpose: this is edited on a phone, held
 * in one hand, next to the box it describes. A twelve-column table on a 390px
 * screen is a horizontal scroll nobody completes.
 *
 * Guessed values are marked. That marking is not decoration — it is what makes
 * a heuristic that reads barcodes by shape (§4.5) safe to ship at all.
 */

function Field({ label, children, guessed, issue }) {
  return (
    <label className={`as-field${issue ? ' as-field-bad' : ''}`}>
      <span className="as-field-label">
        {label}
        {guessed && <span className="as-guess" title="Worked out from the barcode">guessed</span>}
      </span>
      {children}
      {issue && <span className="as-field-issue">{issue}</span>}
    </label>
  );
}

export default function DraftCard({ draft, issues = [], onChange, onRemove, index }) {
  const set = (field) => (event) => onChange(setDraftField(draft, field, event.target.value));
  const guessed = (field) => draft.guessed?.includes(field);
  const issueFor = (field) => issues.find((issue) => issue.field === field)?.message;

  const blocking = issues.some((issue) => issue.blocking);
  const codes = draft.additionalCodes ?? [];

  return (
    <article className={`as-draft${blocking ? ' as-draft-blocked' : ''}`}>
      <header className="as-draft-head">
        <span className="as-draft-index">{index}</span>
        <strong className="as-draft-name">
          {[draft.manufacturer, draft.model].filter(Boolean).join(' ')
            || draft.serialNumber
            || 'New item'}
        </strong>
        {onRemove && (
          <button
            type="button"
            className="as-iconbtn"
            onClick={onRemove}
            aria-label="Remove this row"
          >
            <Trash2 size={14} />
          </button>
        )}
      </header>

      {blocking && (
        <p className="as-draft-block">
          <AlertTriangle size={14} />
          {issues.find((issue) => issue.blocking).message}
        </p>
      )}

      <div className="as-draft-body">
        <PhotoInput
          photoId={draft.photoId}
          onChange={(photoId) => onChange({ ...draft, photoId })}
          label="Photo of the item"
        />

        <div className="as-draft-fields">
          <Field label="Category" issue={issueFor('category')}>
            <select value={draft.category} onChange={set('category')}>
              {CATEGORIES.map((category) => (
                <option key={category} value={category}>{category}</option>
              ))}
            </select>
          </Field>

          <Field label="Counted as">
            <select value={draft.trackingMode} onChange={set('trackingMode')}>
              <option value={TRACKED}>One unit, with a serial</option>
              <option value={BULK}>Bulk, by quantity</option>
            </select>
          </Field>

          <Field label="Make">
            <input value={draft.manufacturer} onChange={set('manufacturer')} placeholder="Dell" />
          </Field>

          <Field label="Model" issue={issueFor('model')}>
            <input value={draft.model} onChange={set('model')} placeholder="Latitude 5540" />
          </Field>

          {/* On a line counted by quantity these four describe the box in
              front of you, not the line — they are saved against ITEM 1 of it,
              and the next box of the same thing becomes item 2. Said out loud,
              because typing a serial into a row that reads "× 20" otherwise
              looks like claiming it for all twenty. */}
          {draft.trackingMode === BULK && (
            <p className="as-draft-note">
              The serial, part number and label below belong to this one box —
              they are kept against the individual item, not the whole line.
            </p>
          )}

          <Field
            label="Serial number"
            guessed={guessed('serialNumber')}
            issue={issueFor('serialNumber')}
          >
            <input value={draft.serialNumber} onChange={set('serialNumber')} />
          </Field>

          <Field label="Part number" guessed={guessed('partNumber')}>
            <input value={draft.partNumber} onChange={set('partNumber')} />
          </Field>

          {/* Offered whenever both codes are present, because the guess is
              wrong often enough that retyping two barcodes by hand is a
              correction people would rather skip than make. */}
          {draft.serialNumber && draft.partNumber && (
            <button
              type="button"
              className="as-swap"
              onClick={() => onChange(swapSerialAndPart(draft))}
            >
              <RefreshCw size={13} />
              The other way round — this is the part number, that is the serial
            </button>
          )}

          <Field label="Asset label" issue={issueFor('assetTag')}>
            <input value={draft.assetTag} onChange={set('assetTag')} placeholder="PMW-0142" />
          </Field>

          {draft.trackingMode === BULK && (
            <Field label="Quantity" issue={issueFor('quantity')}>
              <input
                type="number"
                min="1"
                inputMode="numeric"
                value={draft.quantity}
                onChange={set('quantity')}
              />
            </Field>
          )}

          {/* A condition describes a thing, and a bulk line is a count of
              things. "All new" typed here would be written against item 1
              alone — twenty new cables recorded as one new cable and nineteen
              nobody looked at — so it is asked per item on the row instead. */}
          {draft.trackingMode !== BULK && (
            <Field label="Condition">
              <select value={draft.condition} onChange={set('condition')}>
                {CONDITIONS.map((condition) => (
                  <option key={condition} value={condition}>{condition}</option>
                ))}
              </select>
            </Field>
          )}

          <Field label="Where it is">
            <input value={draft.location} onChange={set('location')} placeholder="Store room" />
          </Field>

          <Field label="Specification">
            <textarea
              rows={2}
              value={draft.specSummary}
              onChange={set('specSummary')}
              placeholder="16GB RAM, 512GB SSD, i7-1355U"
            />
          </Field>

          <Field label="Remarks">
            <textarea rows={2} value={draft.remarks} onChange={set('remarks')} />
          </Field>
        </div>
      </div>

      {codes.length > 0 && (
        <footer className="as-draft-codes">
          <Barcode size={13} />
          {/* Kept verbatim rather than dropped: a code nobody could place is
              still the only copy of what was printed on the box. */}
          <span>Also read: {codes.join(', ')}</span>
        </footer>
      )}
    </article>
  );
}

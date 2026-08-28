import { createPortal } from 'react-dom';
import { useState } from 'react';
import {
  X, ScanLine, AlertTriangle, Check, Plus,
} from '../../../components/ui/Icons';
import Button from '../../../components/ui/Button';
import { useTextScanner, SCAN_STATE } from '../scan/useTextScanner';
import { candidates, SCAN_FIELDS } from '../scan/textScan';
import { labelFor } from '../scan/fieldLabels';
import { useScrollLock } from '../../../hooks/useScrollLock';

/**
 * The camera, over the form, reading the label — and ASKING.
 *
 * It used to fill six fields in and close itself the moment the words stopped
 * changing. That is the wrong trade for printed text: a barcode carries a
 * checksum and either decodes or does not, while a hand-held camera on a
 * sticker reads `8` as `B` and `0` as `O` in one frame out of several, and no
 * part of the answer says which frame that was. Writing that straight into the
 * register makes the camera something to be undone rather than used.
 *
 * So everything it reads is MARKED and nothing is taken. Each settled value
 * gets a tick and a cross; crossing one out is remembered, or the camera —
 * still pointed at the same label — would offer it again half a second later.
 * Which field it goes in is a dropdown, not a verdict: the reader's best guess
 * at what a number IS gets it wrong often enough — a part number read as a
 * serial — that being told "guessed" and then having to accept it anyway was
 * only half an answer. Move it, or cross it out.
 * Writing it recognised but could not name is offered at the bottom to be
 * filed by hand, because a code nobody can place is still the only copy of
 * what was printed on the box.
 *
 * It keeps reading while all this is going on. Nothing is written by itself
 * any more, so another pass costs only battery.
 */

const MESSAGES = {
  [SCAN_STATE.STARTING]: 'Starting the camera…',
  [SCAN_STATE.READING]: 'Hold the camera over the label. Tick anything you want to keep.',
  [SCAN_STATE.DENIED]: 'The camera was blocked. Allow it in the browser, or type the value in.',
  [SCAN_STATE.UNAVAILABLE]: 'No camera on this device. Type the value in instead.',
  [SCAN_STATE.NO_READER]: 'This browser cannot read text from a picture. Type the value in instead.',
};

/** One line of writing the reader could not name, and the fields it could go in. */
function LooseLine({ value, onFile, onDismiss }) {
  const [field, setField] = useState(SCAN_FIELDS[0]);

  return (
    <li className="as-loose">
      <span className="as-sheet-found-value">{value}</span>
      <select value={field} onChange={(event) => setField(event.target.value)}>
        {SCAN_FIELDS.map((name) => (
          <option key={name} value={name}>{labelFor(name)}</option>
        ))}
      </select>
      <button
        type="button"
        className="as-iconbtn"
        onClick={() => onFile(field, value)}
        aria-label={`Put ${value} in ${labelFor(field)}`}
      >
        <Plus size={13} />
      </button>
      <button type="button" className="as-iconbtn" onClick={() => onDismiss(value)} aria-label="Discard">
        <X size={13} />
      </button>
    </li>
  );
}

export default function TextScanSheet({ title = 'Scan the label', onCancel, onUse }) {
  const {
    videoRef, state, error, scan, finish, reject, dismiss,
  } = useTextScanner({ active: true });

  useScrollLock(true);

  const found = candidates(scan);
  const extras = scan.additional ?? [];

  // Where each reading is being filed, when that is not where the reader put
  // it. Keyed by the field it was READ as, which is the name the scan knows it
  // by and the one it has to be crossed off under.
  const [filedAs, setFiledAs] = useState({});
  const fieldOf = (entry) => filedAs[entry.field] ?? entry.field;

  /** One value, taken. The sheet stays open: a label carries several. */
  const take = (entry) => {
    onUse({ [fieldOf(entry)]: entry.value }, scan.guessed, []);
    reject(entry.field);
  };

  const takeAll = () => {
    const values = Object.fromEntries(found.map((entry) => [fieldOf(entry), entry.value]));
    onUse(values, scan.guessed, extras);
    onCancel();
  };

  const broken = state === SCAN_STATE.DENIED
    || state === SCAN_STATE.UNAVAILABLE
    || state === SCAN_STATE.NO_READER;
  const failed = state === SCAN_STATE.DONE && !found.length && !extras.length;

  // Hung off the body rather than left where it was opened: the camera covers
  // the SCREEN, and inside the page it would be positioned against whatever
  // the shell's entrance animation left transformed -- which is the page, not
  // the screen, and puts the viewfinder wherever the reader is not.
  return createPortal(
    <div className="as-sheet" role="dialog" aria-modal="true" aria-label={title}>
      <div className="as-sheet-inner">
        <header className="as-sheet-head">
          <ScanLine size={16} />
          <strong>{title}</strong>
          <button type="button" className="as-iconbtn" onClick={onCancel} aria-label="Close">
            <X size={16} />
          </button>
        </header>

        {!broken && (
          <div className="as-sheet-view">
            <video ref={videoRef} className="as-sheet-video" playsInline muted />
            <div className="as-sheet-frame" aria-hidden="true" />
          </div>
        )}

        <p className={`as-sheet-note${broken || failed ? ' as-sheet-note-bad' : ''}`}>
          {(broken || failed) && <AlertTriangle size={14} />}
          {failed
            ? 'Nothing readable on that label. Move closer, hold it still, or type the value in.'
            : MESSAGES[state] ?? MESSAGES[SCAN_STATE.READING]}
        </p>

        {error && broken && <p className="as-sheet-note as-sheet-note-bad">{error}</p>}

        {found.length > 0 && (
          <ul className="as-sheet-found as-sheet-offer">
            {found.map((entry) => (
              <li key={entry.field}>
                <span className="as-sheet-found-value">{entry.value}</span>
                {/* The field is a choice, not a verdict. It opens on what the
                    reader decided, and says so when that was a guess. */}
                <select
                  className="as-sheet-found-field"
                  value={fieldOf(entry)}
                  onChange={(event) => setFiledAs(
                    (current) => ({ ...current, [entry.field]: event.target.value }),
                  )}
                  aria-label={`Which field ${entry.value} goes in`}
                >
                  {SCAN_FIELDS.map((name) => (
                    <option key={name} value={name}>{labelFor(name)}</option>
                  ))}
                </select>
                {/* Said out loud, because a guess that cannot be spotted is
                    a guess that gets saved. */}
                {entry.guessed && (
                  <span className="as-guess" title="Worked out from the shape of the writing">
                    guessed
                  </span>
                )}
                <button
                  type="button"
                  className="as-iconbtn as-take"
                  onClick={() => take(entry)}
                  aria-label={`Use ${entry.value} as ${labelFor(fieldOf(entry))}`}
                >
                  <Check size={13} />
                </button>
                <button
                  type="button"
                  className="as-iconbtn"
                  onClick={() => reject(entry.field)}
                  aria-label={`Discard ${entry.value}`}
                >
                  <X size={13} />
                </button>
              </li>
            ))}
          </ul>
        )}

        {extras.length > 0 && (
          <>
            <p className="as-sheet-note">Also on the label:</p>
            <ul className="as-sheet-found">
              {extras.map((value) => (
                <LooseLine
                  key={value}
                  value={value}
                  onFile={(field, text) => { onUse({ [field]: text }, [], []); dismiss(text); }}
                  onDismiss={dismiss}
                />
              ))}
            </ul>
          </>
        )}

        <footer className="as-sheet-foot">
          <Button variant="secondary" size="sm" onClick={onCancel}>Close</Button>
          {found.length > 0 && (
            <Button size="sm" onClick={takeAll}>
              Take all {found.length}
            </Button>
          )}
          {/* The way out of a label it cannot settle on, while it is still
              trying and has nothing to show for it. */}
          {state === SCAN_STATE.READING && !found.length && (
            <Button variant="secondary" size="sm" onClick={finish}>Stop reading</Button>
          )}
        </footer>
      </div>
    </div>,
    document.body,
  );
}

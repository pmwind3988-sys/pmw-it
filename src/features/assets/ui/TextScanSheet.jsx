import { useEffect } from 'react';
import { X, ScanLine, AlertTriangle, Check } from '../../../components/ui/Icons';
import Button from '../../../components/ui/Button';
import { useTextScanner, SCAN_STATE } from '../scan/useTextScanner';
import { settledValues } from '../scan/textScan';
import { labelFor } from '../scan/fieldLabels';
import { useScrollLock } from '../../../hooks/useScrollLock';

/**
 * The camera, over the form, reading the label.
 *
 * It fills in and closes itself once the words stop changing (§ the
 * settling rule in `textScan.js`). Waiting for a person to press an
 * "accept" button would mean holding a phone steady over a box with one
 * hand while reading a list with the other.
 *
 * What it has read so far is on screen the whole time, because a
 * recogniser that works silently for four seconds and then changes six
 * fields is one nobody trusts the second time.
 */

const MESSAGES = {
  [SCAN_STATE.STARTING]: 'Starting the camera…',
  [SCAN_STATE.READING]: 'Hold the camera over the label until it stops changing.',
  [SCAN_STATE.DENIED]: 'The camera was blocked. Allow it in the browser, or type the value in.',
  [SCAN_STATE.UNAVAILABLE]: 'No camera on this device. Type the value in instead.',
  [SCAN_STATE.NO_READER]: 'This browser cannot read text from a picture. Type the value in instead.',
};

export default function TextScanSheet({ title = 'Scan the label', onCancel, onUse }) {
  const { videoRef, state, error, scan, finish } = useTextScanner({ active: true });

  // The page behind a full-screen camera must not scroll under the picture.
  useScrollLock(true);

  const found = settledValues(scan);
  const names = Object.keys(found);

  // Filling the form is the end of the scan, so it happens here rather
  // than being offered: the phone is in the air over a box, and the next
  // thing wanted is the form, not a confirmation.
  useEffect(() => {
    if (state !== SCAN_STATE.DONE || !names.length) return;
    onUse(found, scan.guessed, scan.additional);
  }, [state, names.length]); // eslint-disable-line react-hooks/exhaustive-deps

  const failed = state === SCAN_STATE.DONE && !names.length;
  const broken = state === SCAN_STATE.DENIED
    || state === SCAN_STATE.UNAVAILABLE
    || state === SCAN_STATE.NO_READER;

  return (
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

        {names.length > 0 && (
          <ul className="as-sheet-found">
            {names.map((field) => (
              <li key={field}>
                <Check size={13} />
                <span className="as-sheet-found-label">{labelFor(field)}</span>
                <span className="as-sheet-found-value">{found[field]}</span>
              </li>
            ))}
          </ul>
        )}

        <footer className="as-sheet-foot">
          <Button variant="secondary" size="sm" onClick={onCancel}>Cancel</Button>
          {/* The way out of a label it cannot settle on: take the fields it
              did read rather than losing them to a scan that never ends. */}
          {state === SCAN_STATE.READING && (
            <Button
              size="sm"
              onClick={names.length ? () => onUse(found, scan.guessed, scan.additional) : finish}
            >
              {names.length ? 'Use what it has' : 'Stop reading'}
            </Button>
          )}
        </footer>
      </div>
    </div>
  );
}

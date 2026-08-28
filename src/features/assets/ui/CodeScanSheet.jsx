import { createPortal } from 'react-dom';
import { useCallback, useMemo, useState } from 'react';
import { X, ScanLine, AlertTriangle, Check, Barcode } from '../../../components/ui/Icons';
import Button from '../../../components/ui/Button';
import { useScanner, CAMERA_STATE } from '../scan/useScanner';
import { classifyCodes } from '../scan/classifyCode';
import { labelFor } from '../scan/fieldLabels';
import ScanControls from './ScanControls';
import { useScrollLock } from '../../../hooks/useScrollLock';

/**
 * The camera, over one box, reading its barcodes.
 *
 * Different from `TextScanSheet` in the one way that matters: that one reads
 * printed words and stops when they stop changing, this one reads barcodes and
 * WAITS. A box carries several codes — the serial on one sticker, the part
 * number on another — and which of them is which can only be answered by
 * looking at all of them together (`classifyCode.js`). Closing after the first
 * one would file whichever sticker happened to be nearest as the serial number.
 *
 * So it collects, shows what it has worked out as it goes, and hands over when
 * the person holding the phone says the box is done. Nothing is written into
 * the form until then, which is what makes pointing it at the wrong box
 * harmless.
 */

const MESSAGES = {
  [CAMERA_STATE.STARTING]: 'Starting the camera…',
  [CAMERA_STATE.RUNNING]: 'Point at each barcode on the box. Take your time — it waits.',
  [CAMERA_STATE.DENIED]: 'The camera was blocked. Allow it in the browser, or type the code in.',
  [CAMERA_STATE.UNAVAILABLE]: 'No camera on this device. Type the code in instead.',
  [CAMERA_STATE.NO_DECODER]: 'This browser cannot read barcodes. Type the code in instead.',
};

/** Which fields a scan may fill, in the order they are shown back. */
const FILLED = ['serialNumber', 'partNumber', 'macAddress', 'assetTag'];

export default function CodeScanSheet({ title = 'Scan the barcodes', onCancel, onUse }) {
  const [codes, setCodes] = useState([]);

  // Held by value rather than by index, because the same barcode stays in
  // frame and is decoded eight times a second. Counting each of those would
  // turn one sticker into a serial number, a part number and four leftovers.
  const onCodes = useCallback((found) => {
    setCodes((seen) => {
      const known = new Set(seen.map((code) => code.rawValue));
      const added = [];
      for (const code of found) {
        const raw = String(code?.rawValue ?? '').trim();
        if (!raw || known.has(raw)) continue;
        known.add(raw);
        added.push({ rawValue: raw, format: code.format ?? '' });
      }
      return added.length ? [...seen, ...added] : seen;
    });
  }, []);

  const {
    videoRef, state, error, controls, torchOn, toggleTorch, zoomTo, focusOn, quiet,
  } = useScanner({ active: true, onCodes });

  // The page behind a full-screen camera must not scroll under the picture.
  useScrollLock(true);

  const tapToFocus = (event) => {
    const box = event.currentTarget.getBoundingClientRect();
    if (!box.width || !box.height) return;
    focusOn((event.clientX - box.left) / box.width, (event.clientY - box.top) / box.height);
  };

  // Re-run on every new code, so the labels under the viewfinder change as the
  // second sticker arrives — which is exactly when the guess changes from "the
  // only code is the serial" to "the shared one is the part number".
  const reading = useMemo(() => classifyCodes(codes), [codes]);
  const named = FILLED.filter((field) => reading[field]);

  const broken = state === CAMERA_STATE.DENIED
    || state === CAMERA_STATE.UNAVAILABLE
    || state === CAMERA_STATE.NO_DECODER;

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
          <div className="as-sheet-view" onClick={tapToFocus}>
            <video ref={videoRef} className="as-sheet-video" playsInline muted />
            <div className="as-sheet-frame" aria-hidden="true" />

            <ScanControls
              controls={controls}
              torchOn={torchOn}
              onTorch={toggleTorch}
              onZoom={zoomTo}
            />

            {quiet && state === CAMERA_STATE.RUNNING && (
              <p className="as-camera-hint">
                Nothing read yet. Fill the white box with the barcode and move a little
                closer{controls.torch ? ', or turn the light on' : ''}.
              </p>
            )}
          </div>
        )}

        <p className={`as-sheet-note${broken ? ' as-sheet-note-bad' : ''}`}>
          {broken && <AlertTriangle size={14} />}
          {MESSAGES[state] ?? MESSAGES[CAMERA_STATE.RUNNING]}
        </p>

        {error && broken && <p className="as-sheet-note as-sheet-note-bad">{error}</p>}

        {named.length > 0 && (
          <ul className="as-sheet-found">
            {named.map((field) => (
              <li key={field}>
                <Check size={13} />
                <span className="as-sheet-found-label">
                  {labelFor(field)}
                  {/* Said out loud, because a guess that cannot be spotted is
                      a guess that gets saved. */}
                  {reading.guessed.includes(field) && (
                    <span className="as-guess" title="Worked out from the shape of the code">
                      guessed
                    </span>
                  )}
                </span>
                <span className="as-sheet-found-value">{reading[field]}</span>
              </li>
            ))}
          </ul>
        )}

        {codes.length > 0 && named.length === 0 && (
          <p className="as-sheet-note">
            <Barcode size={13} /> Read {codes.length}, still working out which is which.
          </p>
        )}

        <footer className="as-sheet-foot">
          <Button variant="secondary" size="sm" onClick={onCancel}>Cancel</Button>
          <Button size="sm" disabled={!named.length} onClick={() => onUse(reading)}>
            {named.length ? `Use ${named.length === 1 ? 'this' : 'these'}` : 'Nothing read yet'}
          </Button>
        </footer>
      </div>
    </div>,
    document.body,
  );
}

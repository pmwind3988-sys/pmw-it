import { useCallback, useEffect, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { useScanner, CAMERA_STATE } from '../../scan/useScanner';
import ScanControls from '../../ui/ScanControls';
import { AlertTriangle, Check, Camera, Plus, X } from '../../../../components/ui/Icons';
import { newPicks, seePicks, takePick, visiblePicks } from '../pickList';

/**
 * The till's camera, over the whole screen.
 *
 * It takes nothing on its own. Every barcode it reads is listed on a card over
 * the picture (`till/pickList.js` decides which), and only a code somebody
 * TAPS goes on the list — a box carries four or five barcodes and the person
 * holding it is the one who knows which they want. A tapped code reads
 * "added" while it stays in view and cannot be tapped twice; one the list
 * already holds (`skip`) is never offered.
 *
 * Hung off the body, like the till's sheets, because the shell leaves the
 * page transformed and a fixed box inside it would be placed against the
 * page. It sits UNDER the sheets (a question a tapped code raises opens over
 * it), unless `over` — the serial run's camera, which belongs above its sheet.
 */
export default function TillCamera({
  active, onTake, skip, describe, flash, onQuiet, onClose, status = '', over = false,
}) {
  const [picks, setPicks] = useState(newPicks);
  const [now, setNow] = useState(() => Date.now());
  const picksRef = useRef(picks);
  const skipRef = useRef(skip);
  useEffect(() => { skipRef.current = skip; }, [skip]);

  const onCodes = useCallback((codes) => {
    const at = Date.now();
    const next = seePicks(picksRef.current, codes, { now: at, skip: (code) => Boolean(skipRef.current?.(code)) });
    picksRef.current = next;
    setPicks(next);
    setNow(at);
  }, []);

  const {
    videoRef, state, controls, torchOn, toggleTorch, zoomTo, focusOn, quiet,
  } = useScanner({ active, onCodes });

  const shown = visiblePicks(picks, now);

  // Frames only arrive while something is being read, so the card is cleared
  // by the clock: without it the last box's codes would sit there for good.
  const showing = shown.length > 0;
  useEffect(() => {
    if (!showing) return undefined;
    const timer = setInterval(() => setNow(Date.now()), 500);
    return () => clearInterval(timer);
  }, [showing]);

  // A camera that has been looking at something for a while without reading
  // it is the moment to offer the other ways in — not before.
  useEffect(() => {
    onQuiet?.(Boolean(quiet && state === CAMERA_STATE.RUNNING));
  }, [quiet, state, onQuiet]);

  // The page under a full-screen camera must not scroll behind it.
  useEffect(() => {
    const before = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => { document.body.style.overflow = before; };
  }, []);

  // Escape closes the camera — only while it is the thing on top. With a
  // sheet open over it, Escape belongs to the sheet.
  useEffect(() => {
    if (!active) return undefined;
    const onKey = (event) => { if (event.key === 'Escape') onClose?.(); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [active, onClose]);

  const tap = (code) => {
    const next = takePick(picksRef.current, code);
    picksRef.current = next;
    setPicks(next);
    onTake(code);
  };

  const tapToFocus = (event) => {
    const box = event.currentTarget.getBoundingClientRect();
    if (!box.width || !box.height) return;
    focusOn((event.clientX - box.left) / box.width, (event.clientY - box.top) / box.height);
  };

  const stop = (event) => event.stopPropagation();

  return createPortal(
    <div className={over ? 'till-fullcam till-fullcam-over' : 'till-fullcam'} role="dialog" aria-modal="true" aria-label="Camera">
      <div className="till-viewfinder as-viewfinder" onClick={tapToFocus}>
        <video ref={videoRef} playsInline muted className="as-video" />
        <div className="as-reticle" aria-hidden="true" />
        <ScanControls controls={controls} torchOn={torchOn} onTorch={toggleTorch} onZoom={zoomTo} />

        <div className="till-fullcam-bar" onClick={stop}>
          <button type="button" className="till-fullcam-close" onClick={onClose} aria-label="Close camera">
            <X size={20} />
          </button>
          {status && <span className="till-fullcam-status till-mono">{status}</span>}
        </div>

        {flash && (
          <div className={`till-flash till-fullcam-flash till-flash-${flash.kind}`} role="status">
            {flash.kind === 'ok' ? <Check size={15} /> : <AlertTriangle size={15} />}
            <span>{flash.text}</span>
          </div>
        )}

        {state === CAMERA_STATE.STARTING && <p className="as-camera-msg">Starting the camera…</p>}
        {state === CAMERA_STATE.DENIED && (
          <div className="as-camera-msg">
            <Camera size={22} />
            <p>This browser is not allowing the camera. Close it and type the code instead.</p>
          </div>
        )}
        {(state === CAMERA_STATE.UNAVAILABLE || state === CAMERA_STATE.NO_DECODER) && (
          <div className="as-camera-msg">
            <Camera size={22} />
            <p>No camera that can read barcodes here. Close it and type the code instead.</p>
          </div>
        )}

        {state === CAMERA_STATE.RUNNING && (
          <div className="till-picks" onClick={stop} aria-live="polite">
            <strong className="till-picks-head">{shown.length ? 'Seen on camera — tap to add' : 'Point at a barcode'}</strong>
            {shown.map((pick) => (
              <button
                key={pick.code}
                type="button"
                className={pick.taken ? 'till-pick taken' : 'till-pick'}
                disabled={pick.taken}
                onClick={() => tap(pick.code)}
              >
                <span className="till-pick-text">
                  <span className="till-mono">{pick.code}</span>
                  <span>{pick.taken ? 'Added' : (describe?.(pick.code) ?? '')}</span>
                </span>
                <span className="till-pick-mark" aria-hidden="true">{pick.taken ? <Check size={18} /> : <Plus size={18} />}</span>
              </button>
            ))}
          </div>
        )}
      </div>
    </div>,
    document.body,
  );
}

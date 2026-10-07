import { useEffect } from 'react';
import { useScanner, CAMERA_STATE } from '../../scan/useScanner';
import ScanControls from '../../ui/ScanControls';
import { AlertTriangle, Check, Camera } from '../../../../components/ui/Icons';

/**
 * The till's viewfinder. Inline, not over the whole screen: the receipt under
 * it IS the feedback, and a camera that hides the list hides whether the last
 * box counted.
 */
export default function TillCamera({
  active, onCodes, flash, onQuiet, choices = [], onPick, onDismiss,
}) {
  const {
    videoRef, state, controls, torchOn, toggleTorch, zoomTo, focusOn, quiet,
  } = useScanner({ active, onCodes });

  // A camera that has been looking at something for a while without reading
  // it is the moment to offer the other ways in — not before.
  useEffect(() => {
    onQuiet?.(Boolean(quiet && state === CAMERA_STATE.RUNNING));
  }, [quiet, state, onQuiet]);

  const tapToFocus = (event) => {
    const box = event.currentTarget.getBoundingClientRect();
    if (!box.width || !box.height) return;
    focusOn((event.clientX - box.left) / box.width, (event.clientY - box.top) / box.height);
  };

  return (
    <div className="till-viewfinder as-viewfinder" onClick={tapToFocus}>
      <video ref={videoRef} playsInline muted className="as-video" />
      <div className="as-reticle" aria-hidden="true" />
      <ScanControls controls={controls} torchOn={torchOn} onTorch={toggleTorch} onZoom={zoomTo} />

      {/* Several barcodes in the aiming box and no way to tell which is
          meant: the person holding the box picks. Clicks stop here so a tap
          on a choice is not also a tap-to-focus. */}
      {choices.length > 0 && (
        <div className="till-choose" role="dialog" aria-label="Which barcode?" onClick={(event) => event.stopPropagation()}>
          <div className="till-choose-head">
            <strong>Which one did you mean?</strong>
            <button type="button" className="till-link" onClick={onDismiss}>Neither</button>
          </div>
          {choices.map((choice) => (
            <button key={choice.code} type="button" className="till-choice" onClick={() => onPick(choice.code)}>
              <span className="till-mono">{choice.code}</span>
              <span>{choice.kind}</span>
            </button>
          ))}
        </div>
      )}

      {flash && !choices.length && (
        <div className={`till-flash till-flash-${flash.kind}`} role="status">
          {flash.kind === 'ok' ? <Check size={15} /> : <AlertTriangle size={15} />}
          <span>{flash.text}</span>
        </div>
      )}

      {state === CAMERA_STATE.STARTING && <p className="as-camera-msg">Starting the camera…</p>}
      {state === CAMERA_STATE.DENIED && (
        <div className="as-camera-msg">
          <Camera size={22} />
          <p>This browser is not allowing the camera. Type or search in the box below instead.</p>
        </div>
      )}
      {(state === CAMERA_STATE.UNAVAILABLE || state === CAMERA_STATE.NO_DECODER) && (
        <div className="as-camera-msg">
          <Camera size={22} />
          <p>No camera that can read barcodes here. Type or search in the box below.</p>
        </div>
      )}
    </div>
  );
}

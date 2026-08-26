import { Sun } from '../../../components/ui/Icons';

/**
 * The light and the zoom, over the picture.
 *
 * Both are the answer to the same complaint — the camera never reads the
 * barcode — and neither is something the decoder can fix: a store room is dark
 * and a label on a pallet is too far away to fill the frame.
 *
 * Both are OPTIONAL. Support varies by handset, so each control appears only
 * where the camera reported it (`cameraTrack.js`); a phone with neither shows
 * nothing here rather than two dead buttons.
 */

export default function ScanControls({ controls, torchOn, onTorch, onZoom }) {
  if (!controls?.torch && !controls?.zoom) return null;

  return (
    // A press here is a press on a control, not "focus on that spot" -- the
    // picture behind it listens for taps.
    <div className="as-camctl" onClick={(event) => event.stopPropagation()}>
      {controls.torch && (
        <button
          type="button"
          className={`as-camctl-btn${torchOn ? ' as-camctl-on' : ''}`}
          onClick={onTorch}
          aria-pressed={torchOn}
        >
          <Sun size={16} />
          <span>{torchOn ? 'Light on' : 'Light'}</span>
        </button>
      )}

      {controls.zoom && (
        <label className="as-camctl-zoom">
          <span className="as-camctl-label">Zoom</span>
          <input
            type="range"
            min={controls.zoom.min}
            max={controls.zoom.max}
            step={controls.zoom.step}
            defaultValue={controls.zoom.min}
            onChange={(event) => onZoom(event.target.value)}
            aria-label="Zoom the camera in on the barcode"
          />
        </label>
      )}
    </div>
  );
}

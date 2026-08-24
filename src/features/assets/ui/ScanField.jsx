import { ScanLine } from '../../../components/ui/Icons';

/**
 * A field with the camera on the end of it.
 *
 * The button says "start here", not "fill this box". What comes off a
 * label is several values at once, and the one being pointed at is
 * rarely the only one worth keeping — so the scan fills whatever it
 * recognises and the field the button sits beside is simply the one that
 * prompted it. Anything already typed is left alone.
 */

export default function ScanField({ label, children, onScan }) {
  return (
    <div className="as-scanfield">
      {children}
      <button
        type="button"
        className="as-scanfield-btn"
        onClick={onScan}
        aria-label={`Scan ${label} with the camera`}
        title="Read this off the label with the camera"
      >
        <ScanLine size={15} />
      </button>
    </div>
  );
}

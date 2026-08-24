import { useEffect } from 'react';
import { createPortal } from 'react-dom';
import { X, Download } from '../../../components/ui/Icons';

/**
 * One photograph, filling the screen.
 *
 * A thumbnail is enough to say "yes, that is the right box" and never enough
 * to read the serial number printed on the side of it — which is the actual
 * reason anybody photographs a delivery. Tapping the picture is how that is
 * asked for everywhere else, so it is how it is asked for here.
 *
 * Rendered into `document.body` rather than where it is used. A detail page
 * puts its photographs inside cards that clip and stack, and an overlay born
 * inside one of those is trapped by it — the picture would open at the size of
 * the card it came from.
 */
export default function PhotoLightbox({ src, alt, href, onClose }) {
  // Escape closes it, and the page behind must not scroll while it is up: on a
  // phone a scrolling backdrop reads as the picture itself failing to move.
  useEffect(() => {
    const onKey = (event) => {
      if (event.key === 'Escape') onClose();
    };
    document.addEventListener('keydown', onKey);
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      document.removeEventListener('keydown', onKey);
      document.body.style.overflow = previousOverflow;
    };
  }, [onClose]);

  if (!src) return null;

  return createPortal(
    <div
      className="as-lightbox"
      role="dialog"
      aria-modal="true"
      aria-label={alt || 'Photograph'}
    >
      {/* The backdrop is the close button. Anywhere off the picture dismisses
          it, which is the gesture people already try first. */}
      <button
        type="button"
        className="as-lightbox-back"
        onClick={onClose}
        aria-label="Close the photograph"
      />

      <img src={src} alt={alt || 'Photograph'} className="as-lightbox-img" />

      <div className="as-lightbox-bar">
        {alt && <span className="as-lightbox-name">{alt}</span>}
        {href && (
          <a href={href} target="_blank" rel="noreferrer" className="as-lightbox-open">
            <Download size={14} /> Open the original
          </a>
        )}
        <button
          type="button"
          className="as-lightbox-close"
          onClick={onClose}
          aria-label="Close the photograph"
        >
          <X size={18} />
        </button>
      </div>
    </div>,
    document.body,
  );
}

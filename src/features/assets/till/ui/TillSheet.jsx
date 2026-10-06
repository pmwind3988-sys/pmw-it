import { useEffect, useRef } from 'react';
import { createPortal } from 'react-dom';

/**
 * A sheet over the till: from the bottom on a phone, centred on a desk.
 * Hung off the body, because the shell's entrance animation leaves the page
 * transformed and a fixed box inside it would be placed against the page.
 * Escape closes it, and only it — it is handled on the sheet, not the window.
 */
export default function TillSheet({ title, onClose, children, wide = false }) {
  const boxRef = useRef(null);

  // Into the sheet's first text box if it has one — a person search wants the
  // keyboard at once — else its first control. After the current task, so the
  // press that opened the sheet has finished with focus first.
  useEffect(() => {
    const timer = setTimeout(() => {
      const box = boxRef.current;
      const target = box?.querySelector('input:not([type=checkbox]), textarea')
        ?? box?.querySelector('.till-sheet-head ~ * button, .till-sheet-head ~ * select')
        ?? box?.querySelector('button');
      target?.focus();
    }, 0);
    return () => clearTimeout(timer);
  }, []);

  const onKeyDown = (event) => {
    if (event.key === 'Escape') {
      event.stopPropagation();
      onClose();
    }
  };

  return createPortal(
    <div className="till-scrim" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose(); }}>
      <div
        ref={boxRef}
        className={wide ? 'till-sheet till-sheet-wide' : 'till-sheet'}
        role="dialog"
        aria-modal="true"
        aria-label={title}
        onKeyDown={onKeyDown}
      >
        <header className="till-sheet-head">
          <h2>{title}</h2>
          <button type="button" className="till-link" onClick={onClose}>Cancel</button>
        </header>
        {children}
      </div>
    </div>,
    document.body,
  );
}

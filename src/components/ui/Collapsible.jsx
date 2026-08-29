import { useState } from 'react';
import { ChevronDown } from './Icons';

/**
 * A panel that can be folded away.
 *
 * Filters and long detail forms are worth having and not worth looking at all
 * the time — on a phone the register's five dropdowns fill the screen before a
 * single row of it is visible. Folding one away is a tap, and the tap is
 * REMEMBERED: somebody who works from a phone should not have to fold the same
 * panel away every time they open the page.
 *
 * `summary` is what the header says while it is shut, and it matters more than
 * the fold does. A panel that hides how many filters are on is a page that
 * lies about what it is showing.
 */

const remembered = (id, fallback) => {
  if (!id) return fallback;
  try {
    const stored = window.localStorage.getItem(`fold:${id}`);
    if (stored === 'open') return true;
    if (stored === 'shut') return false;
  } catch {
    // A browser with site data switched off still gets a working panel.
  }
  return fallback;
};

const remember = (id, open) => {
  if (!id) return;
  try {
    window.localStorage.setItem(`fold:${id}`, open ? 'open' : 'shut');
  } catch {
    // Same again: the fold works, it just will not be there next time.
  }
};

export default function Collapsible({
  id,
  title,
  summary,
  defaultOpen = true,
  actions,
  className = '',
  children,
}) {
  const [open, setOpen] = useState(() => remembered(id, defaultOpen));

  const toggle = () => {
    setOpen((current) => {
      remember(id, !current);
      return !current;
    });
  };

  return (
    <div className={`ui-fold ${open ? 'ui-fold-open' : 'ui-fold-shut'} ${className}`.trim()}>
      <div className="ui-fold-head">
        <button
          type="button"
          className="ui-fold-toggle"
          aria-expanded={open}
          onClick={toggle}
        >
          <ChevronDown size={15} className="ui-fold-arrow" />
          <span className="ui-fold-title">{title}</span>
          {summary && <span className="ui-fold-summary">{summary}</span>}
        </button>
        {/* Outside the button on purpose: a control inside a control cannot be
            reached by a keyboard, and a Refresh nested in a toggle would fold
            the panel every time it was pressed. */}
        {actions && <div className="ui-fold-actions">{actions}</div>}
      </div>
      {open && <div className="ui-fold-body">{children}</div>}
    </div>
  );
}

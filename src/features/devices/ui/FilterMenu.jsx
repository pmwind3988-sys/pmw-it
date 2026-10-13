import { Fragment, useEffect, useRef } from 'react';
import { ChevronDown } from '../../../components/ui/Icons';

/**
 * The "+ Filter" button and the menu it opens. The menu is grouped (Place,
 * Health, Security, Hardware & upkeep) and every choice writes one filter.
 *
 * Controlled by the table, so a chip can open it straight at its own group:
 * `focusFilter` names the filter whose first option takes the focus when the
 * menu opens. Escape and a press outside close it and hand the focus back to
 * the button that opened it.
 */
export default function FilterMenu({ open, onOpenChange, focusFilter, groups, onPick }) {
  const wrap = useRef(null);
  const trigger = useRef(null);
  const menu = useRef(null);

  // Escape and outside presses close the menu. The listeners exist only while
  // it is open, so a closed table listens to nothing.
  useEffect(() => {
    if (!open) return undefined;

    const onPointer = (event) => {
      if (wrap.current && !wrap.current.contains(event.target)) onOpenChange(false);
    };
    const onKey = (event) => {
      if (event.key !== 'Escape') return;
      onOpenChange(false);
      trigger.current?.focus();
    };

    document.addEventListener('pointerdown', onPointer);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('pointerdown', onPointer);
      document.removeEventListener('keydown', onKey);
    };
  }, [open, onOpenChange]);

  // Focus moves into the menu when it opens: to the named option when a chip
  // opened it, otherwise to the first choice. Only the DOM is touched here.
  useEffect(() => {
    if (!open || !menu.current) return;
    const target = focusFilter
      ? menu.current.querySelector(`[data-filter="${focusFilter}"]`)
      : menu.current.querySelector('button');
    target?.focus();
  }, [open, focusFilter]);

  return (
    <div className="dtf" ref={wrap}>
      <button
        ref={trigger}
        type="button"
        className="dtf-trigger"
        aria-expanded={open}
        aria-controls="dtf-menu"
        onClick={() => onOpenChange(!open)}
      >
        + Filter
        <ChevronDown size={14} />
      </button>

      {open && (
        <div className="dtf-menu" id="dtf-menu" ref={menu} role="region" aria-label="Filter choices">
          {groups.map((group) => (
            <section className="dtf-group" key={group.id} aria-label={group.title}>
              <h3 className="dtf-title">{group.title}</h3>
              <div className="dtf-items">
                {group.items.map((item, index) => (
                  <Fragment key={`${item.filter}:${item.value}`}>
                  {item.sub && item.sub !== group.items[index - 1]?.sub && (
                    <span className="dtf-sub">{item.sub}</span>
                  )}
                  <button
                    type="button"
                    key={`${item.filter}:${item.value}`}
                    data-filter={item.filter}
                    className="dtf-item"
                    onClick={() => onPick(item.filter, item.value)}
                  >
                    {item.label}
                  </button>
                  </Fragment>
                ))}
              </div>
            </section>
          ))}
        </div>
      )}
    </div>
  );
}

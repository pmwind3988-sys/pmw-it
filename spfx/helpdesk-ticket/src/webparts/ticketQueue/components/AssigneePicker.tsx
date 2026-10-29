// "Assigned to": a button that shows who the list is for, and a checklist behind it.
import * as React from 'react';
import { assigneesOf, UNASSIGNED } from '../queue/queueFilter';
import { PersonRef, QueueTicket } from '../queue/ticket';
import { useInstanceId } from '../../helpdeskTicket/components/hooks';
import { ChevronDownIcon } from './queueIcons';
import styles from './TicketQueue.module.scss';

/** Above this many people, a search box sits at the top of the list. */
const SEARCH_ABOVE = 8;

export interface AssigneePickerProps {
  ids: number[];
  me: PersonRef;
  /** Every ticket, so each name can show how many it holds. */
  rows: QueueTicket[];
  summary: string;
  onChange(ids: number[]): void;
}

export function AssigneePicker(props: AssigneePickerProps): React.ReactElement {
  const { ids, me, rows, summary, onChange } = props;
  const [open, setOpen] = React.useState(false);
  const [term, setTerm] = React.useState('');
  const wrap = React.useRef<HTMLDivElement>(null);
  const button = React.useRef<HTMLButtonElement>(null);
  const panel = React.useRef<HTMLDivElement>(null);
  const panelId = useInstanceId('ap-panel');
  const searchId = useInstanceId('ap-search');

  const others = assigneesOf(rows).filter((p) => p.person.id !== me.id);
  const myCount = rows.filter((t) => t.assignedTo?.id === me.id).length;
  const unassignedCount = rows.filter((t) => !t.assignedTo).length;
  const needle = term.trim().toLowerCase();
  const visibleOthers = others.filter(
    (o) => !needle || `${o.person.name} ${o.person.email}`.toLowerCase().indexOf(needle) !== -1
  );

  // Focus moves into the list when it opens.
  React.useEffect(() => {
    if (open) panel.current?.focus();
  }, [open]);

  // A click anywhere outside the picker closes it.
  React.useEffect(() => {
    if (!open) return undefined;
    const onDown = (e: MouseEvent): void => {
      const target = e.target;
      if (target instanceof Node && wrap.current && !wrap.current.contains(target)) setOpen(false);
    };
    document.addEventListener('mousedown', onDown);
    return () => document.removeEventListener('mousedown', onDown);
  }, [open]);

  const close = (refocus: boolean): void => {
    setOpen(false);
    if (refocus) button.current?.focus();
  };

  const toggle = (id: number): void => {
    onChange(ids.indexOf(id) !== -1 ? ids.filter((x) => x !== id) : [...ids, id]);
  };

  const onPanelKey = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    if (e.key === 'Escape') {
      e.stopPropagation();
      close(true);
    }
  };

  const checked = (id: number): boolean => ids.indexOf(id) !== -1;

  return (
    <div className={styles.pickerWrap} ref={wrap}>
      <button
        ref={button}
        type="button"
        className={styles.pickerBtn}
        aria-haspopup="dialog"
        aria-expanded={open}
        aria-controls={open ? panelId : undefined}
        onClick={() => setOpen((o) => !o)}
      >
        <span className={styles.pickerText}>{summary}</span>
        <ChevronDownIcon />
      </button>

      {open && (
        <div
          id={panelId}
          ref={panel}
          role="dialog"
          aria-label="Choose who the list is for"
          tabIndex={-1}
          className={styles.popover}
          onKeyDown={onPanelKey}
        >
          <div className={styles.quickRow}>
            <button type="button" className={styles.linkBtn} onClick={() => onChange([me.id])}>
              Just me
            </button>
            <button type="button" className={styles.linkBtn} onClick={() => onChange([])}>
              Everyone
            </button>
          </div>

          {others.length > SEARCH_ABOVE && (
            <>
              <label htmlFor={searchId} className={styles.visuallyHidden}>
                Find a person
              </label>
              <input
                id={searchId}
                type="search"
                className={styles.popSearch}
                placeholder="Find a person"
                value={term}
                onChange={(e) => setTerm(e.target.value)}
              />
            </>
          )}

          <ul className={styles.checkList}>
            <li>
              <label className={styles.checkRow}>
                <input type="checkbox" checked={checked(me.id)} onChange={() => toggle(me.id)} />
                <span className={styles.checkName}>{me.name ? `Me (${me.name})` : 'Me'}</span>
                <span className={styles.count}>{myCount}</span>
              </label>
            </li>
            <li>
              <label className={styles.checkRow}>
                <input type="checkbox" checked={checked(UNASSIGNED)} onChange={() => toggle(UNASSIGNED)} />
                <span className={styles.checkName}>Unassigned</span>
                <span className={styles.count}>{unassignedCount}</span>
              </label>
            </li>
            {visibleOthers.map((o) => (
              <li key={o.person.id}>
                <label className={styles.checkRow}>
                  <input type="checkbox" checked={checked(o.person.id)} onChange={() => toggle(o.person.id)} />
                  <span className={styles.checkName}>{o.person.name}</span>
                  <span className={styles.count}>{o.count}</span>
                </label>
              </li>
            ))}
          </ul>
          {others.length > 0 && visibleOthers.length === 0 && <p className={styles.muted}>No one matches.</p>}
          <div className={styles.popFoot}>
            <button type="button" className={styles.linkBtn} onClick={() => close(true)}>
              Done
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

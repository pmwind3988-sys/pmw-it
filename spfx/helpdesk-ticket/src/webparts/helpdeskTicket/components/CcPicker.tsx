// "CC": the Tickets list's Cc column (people emailed about the ticket), a debounced directory search with chips for the picked people.
import * as React from 'react';
import { TicketsApi } from '../sharepoint/ticketsApi';
import { Person } from '../tickets/draft';
import { useInstanceId, useMountedRef } from './hooks';
import { CloseIcon } from './icons';
import styles from './HelpdeskTicket.module.scss';

export interface CcPickerProps {
  api: TicketsApi;
  people: Person[];
  onAdd: (person: Person) => void;
  onRemove: (email: string) => void;
}

const SEARCH_DELAY_MS = 300;
const MIN_CHARS = 2;

const sameEmail = (a: string, b: string): boolean => a.toLowerCase() === b.toLowerCase();

export function CcPicker(props: CcPickerProps): React.ReactElement {
  const { api, people, onAdd, onRemove } = props;
  const [query, setQuery] = React.useState('');
  const [open, setOpen] = React.useState(false);
  const [active, setActive] = React.useState(-1);
  // The last answer, tagged with the text it answers. Only an answer for the
  // text now typed is shown, so a slow earlier search cannot overwrite it.
  const [found, setFound] = React.useState<{ q: string; items: Person[]; error: string }>({
    q: '',
    items: [],
    error: ''
  });
  const seq = React.useRef(0);
  const mounted = useMountedRef();
  const inputRef = React.useRef<HTMLInputElement>(null);
  const inputId = useInstanceId('cc');
  const listId = `${inputId}-list`;
  const term = query.trim();

  React.useEffect(() => {
    if (term.length < MIN_CHARS) return undefined;
    const mine = ++seq.current;
    const timer = window.setTimeout(() => {
      api
        .searchPeople(term)
        .then((items) => {
          if (mounted.current && mine === seq.current) setFound({ q: term, items, error: '' });
        })
        .catch(() => {
          if (mounted.current && mine === seq.current) {
            setFound({ q: term, items: [], error: 'Could not search people just now. Try again in a moment.' });
          }
        });
    }, SEARCH_DELAY_MS);
    return () => window.clearTimeout(timer);
  }, [term, api, mounted]);

  const current = term.length >= MIN_CHARS && found.q === term;
  const items = current ? found.items.filter((p) => !people.some((c) => sameEmail(c.email, p.email))) : [];
  const listShown = open && term.length >= MIN_CHARS;

  const pick = (p: Person): void => {
    onAdd(p);
    setQuery('');
    setOpen(false);
    setActive(-1);
    inputRef.current?.focus();
  };

  const onKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    const n = items.length;
    if (e.key === 'ArrowDown' && listShown && n) {
      e.preventDefault();
      setActive((i) => (i + 1) % n);
    } else if (e.key === 'ArrowUp' && listShown && n) {
      e.preventDefault();
      setActive((i) => (i <= 0 ? n - 1 : i - 1));
    } else if (e.key === 'Enter' && listShown && active >= 0 && items[active]) {
      e.preventDefault();
      pick(items[active]);
    } else if (e.key === 'Escape' && open) {
      e.preventDefault();
      setOpen(false);
      setActive(-1);
    }
  };

  const activeId = listShown && active >= 0 && items[active] ? `${listId}-${active}` : undefined;

  return (
    <div className={styles.field}>
      <label htmlFor={inputId} className={styles.label}>
        CC <span className={styles.optional}>(optional)</span>
      </label>
      <span id={`${inputId}-help`} className={styles.hint}>
        People who should also get emails about this ticket, e.g. your manager.
      </span>
      <div className={styles.ccBox} onClick={() => inputRef.current?.focus()}>
        {people.map((p) => {
          const label = p.name || p.email;
          return (
            <span key={p.email} className={styles.chip}>
              {label}
              <button
                type="button"
                className={styles.chipRemove}
                aria-label={`Remove ${label}`}
                onClick={(e) => {
                  e.stopPropagation();
                  onRemove(p.email);
                }}
              >
                <CloseIcon />
              </button>
            </span>
          );
        })}
        <input
          id={inputId}
          ref={inputRef}
          type="text"
          role="combobox"
          className={styles.ccInput}
          placeholder={people.length ? '' : 'Type a name to add someone'}
          aria-describedby={`${inputId}-help`}
          autoComplete="off"
          aria-autocomplete="list"
          aria-expanded={listShown}
          aria-controls={listId}
          aria-activedescendant={activeId}
          value={query}
          onChange={(e) => {
            setQuery(e.target.value);
            setOpen(true);
            setActive(-1);
          }}
          onFocus={() => setOpen(true)}
          onBlur={() => setOpen(false)}
          onKeyDown={onKeyDown}
        />
      </div>
      {listShown && !current && !found.error && (
        <div className={styles.hint} role="status">
          Searching…
        </div>
      )}
      {listShown && current && found.error && <div className={styles.hint}>{found.error}</div>}
      {listShown && current && !found.error && !items.length && (
        <div className={styles.hint} role="status">
          No one matches “{term}”.
        </div>
      )}
      {listShown && current && items.length > 0 && (
        <ul id={listId} role="listbox" aria-label="People found" className={styles.listbox}>
          {items.map((p, i) => (
            <li
              key={p.email}
              id={`${listId}-${i}`}
              role="option"
              aria-selected={i === active}
              className={i === active ? styles.optionActive : styles.option}
              // Keep focus in the box, so the list does not close before the pick lands.
              onMouseDown={(e) => {
                e.preventDefault();
                pick(p);
              }}
              onMouseEnter={() => setActive(i)}
            >
              <span className={styles.optionName}>{p.name || p.email}</span>
              <span className={styles.optionMail}>{p.email}</span>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

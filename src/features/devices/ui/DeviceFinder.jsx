import { useId, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { Search } from '../../../components/ui/Icons';
import { findMatches, machinesWithProgram } from '../finder';
import { labelOf } from '../deviceFilters';
import { statusOf } from '../lifecycle/status';

/** A program search that matches more than this many machines gets a link to the List. */
const PROGRAM_LINK_OVER = 3;

/** The reason a match is shown with, the matched text in bold. */
function Reason({ reason }) {
  const { label, value, suffix, at, length } = reason;
  const head = `${label} `;
  if (at < 0 || length === 0) return <>{head}{value}{suffix}</>;
  return (
    <>
      {head}{value.slice(0, at)}<strong>{value.slice(at, at + length)}</strong>{value.slice(at + length)}{suffix}
    </>
  );
}

/**
 * The "Find a machine" pill. A combobox over every saved machine: arrow keys
 * move through the matches, Enter opens the highlighted one, Escape clears.
 * Each match says why it matched; a program search that finds many machines
 * links to the List filtered to them.
 */
export default function DeviceFinder({ devices }) {
  const navigate = useNavigate();
  const listId = useId();
  const [query, setQuery] = useState('');
  const [active, setActive] = useState(0);
  const [focused, setFocused] = useState(false);

  const matches = useMemo(() => findMatches(devices, query), [devices, query]);
  const results = useMemo(() => matches.map((match) => match.device), [matches]);
  const programCount = useMemo(
    () => (query.trim().length >= 2 ? machinesWithProgram(devices, query.trim()) : 0),
    [devices, query],
  );
  const index = Math.min(active, Math.max(results.length - 1, 0));
  const open = focused && query.trim().length >= 2;
  const showEmpty = open && results.length === 0;

  // Every machine with the program, Retired and Spare included, under "Everything" so the List shows the same figure.
  const openProgramList = () => {
    const text = query.trim();
    setQuery('');
    setFocused(false);
    navigate(`/devices?${new URLSearchParams({ software: text, status: 'All' })}`);
  };

  const openMachine = (device) => {
    if (device?.id === undefined || device?.id === null) return;
    setQuery('');
    setFocused(false);
    navigate(`/devices/${device.id}`);
  };

  const onKeyDown = (event) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      setQuery('');
      setActive(0);
      return;
    }
    if (!open || results.length === 0) return;
    if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault();
      const step = event.key === 'ArrowDown' ? 1 : -1;
      setActive((index + step + results.length) % results.length);
      return;
    }
    if (event.key === 'Enter') {
      event.preventDefault();
      openMachine(results[index]);
    }
  };

  return (
    <div className="dvp-finder" role="search">
      <label className="dvp-finder-pill">
        <Search size={18} />
        <span className="dvp-sr">Find a machine</span>
        <input
          type="text"
          role="combobox"
          aria-expanded={open}
          aria-controls={listId}
          aria-autocomplete="list"
          aria-activedescendant={open && results[index] ? `${listId}-${index}` : undefined}
          autoComplete="off"
          spellCheck={false}
          placeholder="Find a machine — name, owner, serial or AnyDesk ID"
          value={query}
          onChange={(event) => { setQuery(event.target.value); setActive(0); }}
          onFocus={() => setFocused(true)}
          onBlur={() => setFocused(false)}
          onKeyDown={onKeyDown}
        />
      </label>

      {open && results.length > 0 && (
        <ul id={listId} role="listbox" aria-label="Matching machines" className="dvp-finder-list">
          {matches.map(({ device, reason }, position) => (
            <li
              key={device.id ?? device.computerName}
              id={`${listId}-${position}`}
              role="option"
              aria-selected={position === index}
              className={`dvp-finder-opt${position === index ? ' dvp-finder-opt-active' : ''}`}
              // mousedown, not click: the input blurs first and would close the list.
              onMouseDown={(event) => { event.preventDefault(); openMachine(device); }}
              onMouseEnter={() => setActive(position)}
            >
              <strong>{labelOf(device.computerName)}</strong>
              <span className="dvp-finder-meta">
                {labelOf(device.owner)} · {statusOf(device)}
              </span>
              <span className="dvp-finder-meta dvp-finder-why">
                <Reason reason={reason} />
              </span>
            </li>
          ))}
          {programCount > PROGRAM_LINK_OVER && (
            <li role="none" className="dvp-finder-all">
              <button
                type="button"
                // mousedown, as the rows: the input blurs first and would close the list.
                onMouseDown={(event) => { event.preventDefault(); openProgramList(); }}
              >
                All {programCount} machines with {query.trim()} in the List ›
              </button>
            </li>
          )}
        </ul>
      )}

      {showEmpty && (
        <p id={listId} role="status" className="dvp-finder-list dvp-finder-none">
          No machine matches “{query.trim()}”.
        </p>
      )}
    </div>
  );
}

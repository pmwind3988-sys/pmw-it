import { useState } from 'react';
import { Users, X, Search } from '../../../components/ui/Icons';
import { usePeopleSearch } from '../people/usePeopleSearch';
import { initialsOf } from '../../../utils/initials';

/**
 * Who the items are for.
 *
 * The email is the identity — everything per-person keys on it, never on the
 * display name, which two people will spell two ways (§4.5). So a person typed
 * by hand when the directory is unreachable still has to carry one, and the
 * fallback asks for it rather than accepting a bare name.
 */
export default function PersonPicker({ person, onChange }) {
  const [query, setQuery] = useState('');
  const [manual, setManual] = useState({ name: '', email: '' });
  const { results, searching, error, tooShort } = usePeopleSearch(query);

  if (person) {
    return (
      <div className="as-person as-person-picked">
        <span className="as-avatar">{initialsOf(person.name)}</span>
        <span className="as-person-text">
          <strong>{person.name}</strong>
          <span className="as-sub">{person.email}</span>
        </span>
        <button
          type="button"
          className="as-iconbtn"
          onClick={() => onChange(null)}
          aria-label="Pick somebody else"
        >
          <X size={14} />
        </button>
      </div>
    );
  }

  return (
    <div className="as-personpick">
      <label className="as-field">
        <span className="as-field-label">Who is this for?</span>
        <span className="as-searchbox">
          <Search size={15} />
          <input
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder="Search your company directory…"
            autoComplete="off"
          />
        </span>
      </label>

      {tooShort && <p className="as-hint">Keep typing — three letters at least.</p>}
      {searching && <p className="as-hint">Searching…</p>}

      {results.length > 0 && (
        <ul className="as-people">
          {results.map((entry) => (
            <li key={entry.email}>
              <button type="button" className="as-person" onClick={() => onChange(entry)}>
                <span className="as-avatar">{initialsOf(entry.name)}</span>
                <span className="as-person-text">
                  <strong>{entry.name}</strong>
                  <span className="as-sub">
                    {[entry.title, entry.email].filter(Boolean).join(' · ')}
                  </span>
                </span>
              </button>
            </li>
          ))}
        </ul>
      )}

      {!searching && !tooShort && query.length >= 3 && results.length === 0 && !error && (
        <p className="as-hint">Nobody found. Check the spelling, or type them in below.</p>
      )}

      {/* A directory outage must not be the reason a laptop cannot be handed
          over, so the typed fallback is always available — not only on error. */}
      <details className="as-manual" open={Boolean(error)}>
        <summary>
          <Users size={13} /> Type somebody in instead
        </summary>
        {error && <p className="as-field-issue">{error}</p>}
        <div className="as-form">
          <label className="as-field">
            <span className="as-field-label">Name</span>
            <input
              value={manual.name}
              onChange={(event) => setManual({ ...manual, name: event.target.value })}
            />
          </label>
          <label className="as-field">
            <span className="as-field-label">Work email</span>
            <input
              type="email"
              value={manual.email}
              onChange={(event) => setManual({ ...manual, email: event.target.value })}
              placeholder="amir@pmwgroup.com"
            />
          </label>
        </div>
        <button
          type="button"
          className="ui-btn ui-btn-sm ui-btn-secondary"
          disabled={!manual.email.includes('@')}
          onClick={() => onChange({
            name: manual.name.trim() || manual.email.trim(),
            email: manual.email.trim().toLowerCase(),
            login: '',
          })}
        >
          Use this person
        </button>
      </details>
    </div>
  );
}

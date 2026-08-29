import { useMemo, useState } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import { Users, Clock, ScanLine, Pencil } from '../components/ui/Icons';
import { useHandovers } from '../features/assets/useHandovers';
import { peopleWithItems } from '../features/assets/handover/availability';
import { initialsOf } from '../utils/initials';

/**
 * Everyone holding something.
 *
 * Whoever is overdue comes first, then whoever holds the most — the order in
 * which somebody scanning this list would want to act on it.
 */
export default function AssetPeoplePage() {
  const navigate = useNavigate();
  const { handovers, loading, error, reload } = useHandovers();
  const [query, setQuery] = useState('');

  const people = useMemo(() => {
    const all = peopleWithItems(handovers);
    const term = query.trim().toLowerCase();
    if (!term) return all;
    return all.filter(
      (person) => person.name.toLowerCase().includes(term) || person.email.includes(term),
    );
  }, [handovers, query]);

  const overdue = people.reduce((sum, person) => sum + person.overdue, 0);

  return (
    <AppShell
      title="Who has what"
      subtitle={overdue > 0 ? `${overdue} item${overdue === 1 ? '' : 's'} overdue` : 'Everyone currently holding something'}
      search={{ value: query, onChange: setQuery, placeholder: 'Find a person…' }}
      actions={(
        <Button icon={ScanLine} onClick={() => navigate('/assets/handover')}>Hand over</Button>
      )}
    >
      {error && <ErrorBanner message={error} onRetry={reload} />}
      {loading && <div className="spinner" />}

      {!loading && people.length === 0 && (
        <EmptyState>
          <Users size={20} />
          {handovers.length === 0
            ? 'Nothing has been handed to anybody yet.'
            : 'Nobody matches that.'}
        </EmptyState>
      )}

      <ul className="as-people as-people-grid">
        {people.map((person) => (
          <li key={person.email}>
            <Link to={`/assets/people/${encodeURIComponent(person.email)}`} className="as-person">
              <span className="as-avatar">{initialsOf(person.name)}</span>
              <span className="as-person-text">
                <strong>{person.name}</strong>
                <span className="as-sub">
                  {person.units} item{person.units === 1 ? '' : 's'}
                </span>
              </span>
              {person.overdue > 0 && (
                <span className="as-overdue">
                  <Clock size={12} /> {person.overdue}
                </span>
              )}
            </Link>
            {/* Beside the card rather than inside it: a button within a link
                is a control nobody can reach by keyboard and a tap that lands
                on whichever of the two the browser felt like. A misspelt name
                is usually spotted HERE, staring at the list, so the correction
                starts here too — and opens on the person's own page, where
                what they hold is visible while it is made. */}
            <Link
              to={`/assets/people/${encodeURIComponent(person.email)}?edit=1`}
              className="as-iconbtn as-person-edit"
              aria-label={`Correct ${person.name}'s name or email`}
              title="Correct the name or email"
            >
              <Pencil size={13} />
            </Link>
          </li>
        ))}
      </ul>
    </AppShell>
  );
}

import { useState } from 'react';
import { useSharePointToken } from '../../../hooks/useRequests';
import { SHAREPOINT_SITE_URL } from '../useAssets';
import { addCategory } from '../sharepoint/addCategory';
import { categoryRefusal } from '../categories';

/**
 * The category dropdown, with a way out of it.
 *
 * The list this app ships with covers what IT buys most years, and then
 * somebody buys a projector. Until now the only answer was "Other", which
 * loses the one fact worth keeping about it — and typing "Projector" into a
 * SharePoint choice column is refused, so the option has to be created before
 * it can be chosen.
 *
 * Adding one writes to SharePoint immediately, before anything else on the
 * page is saved. That is deliberate: the category has to exist on the column
 * for the row that uses it to save at all, and doing it here means the failure
 * — if there is one — is about the category, said next to the box it was typed
 * into, rather than a save of the whole item failing later for reasons that
 * read as nothing to do with it.
 */
export default function CategoryField({ value, options, onChange }) {
  const getToken = useSharePointToken();
  const [adding, setAdding] = useState(false);
  const [typed, setTyped] = useState('');
  const [busy, setBusy] = useState(false);
  const [problem, setProblem] = useState('');

  const add = async () => {
    const refusal = categoryRefusal(typed, options);
    if (refusal) {
      setProblem(refusal);
      return;
    }

    setBusy(true);
    setProblem('');
    try {
      const tokenRes = await getToken();
      const { category } = await addCategory({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        name: typed,
      });
      onChange(category);
      setTyped('');
      setAdding(false);
    } catch (failure) {
      setProblem(failure.message || 'That category could not be added');
    } finally {
      setBusy(false);
    }
  };

  return (
    <>
      <select
        value={value}
        onChange={(event) => {
          if (event.target.value === '__add') {
            setAdding(true);
            setProblem('');
            return;
          }
          onChange(event.target.value);
        }}
      >
        <option value="">—</option>
        {options.map((option) => (
          <option key={option} value={option}>{option}</option>
        ))}
        {/* A value nobody can select by accident: choosing it opens the box
            below rather than setting the category to the word "__add". */}
        <option value="__add">Add a category…</option>
      </select>

      {adding && (
        <div className="as-addcat">
          <input
            type="text"
            value={typed}
            autoFocus
            placeholder="Projector"
            aria-label="The new category"
            disabled={busy}
            onChange={(event) => { setTyped(event.target.value); setProblem(''); }}
            onKeyDown={(event) => { if (event.key === 'Enter') add(); }}
          />
          <button type="button" className="ui-btn ui-btn-sm ui-btn-primary" disabled={busy} onClick={add}>
            {busy ? 'Adding…' : 'Add'}
          </button>
          <button
            type="button"
            className="ui-btn ui-btn-sm ui-btn-ghost"
            disabled={busy}
            onClick={() => { setAdding(false); setProblem(''); }}
          >
            Cancel
          </button>
          {problem && <p className="as-addcat-problem">{problem}</p>}
          {!problem && (
            <p className="as-addcat-note">
              It becomes an option for everybody, on this list and on handovers.
            </p>
          )}
        </div>
      )}
    </>
  );
}

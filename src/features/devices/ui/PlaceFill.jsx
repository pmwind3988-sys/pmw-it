import { useId, useState } from 'react';
import Button from '../../../components/ui/Button';
import { Pencil } from '../../../components/ui/Icons';

// Owner is corrected, never filled: a machine with nobody on it gets somebody
// through Change owner / Assign, which is what starts an owner-history entry.
const OPENABLE = {
  owner: {
    fields: [{ key: 'owner', label: 'Owner' }],
    ask: 'Fix the spelling of the owner',
  },
  place: {
    fields: [
      { key: 'location', label: 'Location', placeholder: 'e.g. F1' },
      { key: 'department', label: 'Department' },
    ],
    ask: 'Fix the location or department',
  },
};

const textOf = (value) => String(value ?? '').trim();

/**
 * The owner, location and department of a machine, with a pencil beside each
 * group. Nothing is drawn as a form until a pencil asks for it. Saving is a
 * correction of the record -- a misspelt name, a wrong site -- written as a
 * manual edit so the next import leaves it alone. A machine actually going to
 * somebody else still goes through Change owner.
 */
export default function PlaceFill({
  device, owners, locations, departments, onSave, busy,
}) {
  const id = useId();
  const [opened, setOpened] = useState(null);
  const [values, setValues] = useState({});
  // What was just saved, shown until the register's re-read catches up. The
  // page keys this component on the stored values, so the re-read resets it.
  const [saved, setSaved] = useState({});
  const [note, setNote] = useState('');

  const current = (key) => (key in saved ? saved[key] : textOf(device[key]));
  const valueOf = (key) => (key in values ? values[key] : current(key));
  const group = opened ? OPENABLE[opened] : null;
  const options = { owner: owners, location: locations, department: departments };

  const edits = group ? Object.fromEntries(group.fields
    .map(({ key }) => [key, textOf(valueOf(key))])
    // Locations are stored upper-case, so 'f2' over 'F2' is no change; a
    // department's capitalisation is the department's to correct.
    .filter(([key, value]) => value && (key === 'location'
      ? value.toUpperCase() !== current(key).toUpperCase()
      : value !== current(key)))) : {};

  const open = (key) => {
    setNote('');
    setValues({});
    setOpened(key);
  };
  const close = () => {
    setOpened(null);
    setValues({});
  };
  const set = (key) => (event) => setValues((list) => ({ ...list, [key]: event.target.value }));

  const submit = async (event) => {
    event.preventDefault();
    if (!Object.keys(edits).length) return;
    if (await onSave(edits)) {
      const now = { ...edits };
      if ('location' in now) now.location = now.location.toUpperCase();
      setSaved((list) => ({ ...list, ...now }));
      close();
      setNote('Saved.');
    }
  };

  const location = current('location');
  const department = current('department');

  return (
    <>
      <div className="mh-def">
        <span className="mh-def-k">With</span>
        {device.owner ? (
          <>
            <span className="mh-def-v mh-strong">{current('owner') || device.owner}</span>
            <button type="button" className="mh-pencil" aria-label={OPENABLE.owner.ask}
              title={OPENABLE.owner.ask} onClick={() => open('owner')}>
              <Pencil size={14} />
            </button>
          </>
        ) : <span className="mh-def-v mh-muted">Nobody has it</span>}
      </div>

      <div className="mh-def">
        <span className="mh-def-k">At</span>
        <span className="mh-def-v mh-strong">
          {location ? location : (
            <button type="button" className="mh-add" onClick={() => open('place')}>Add location</button>
          )}
          {' · '}
          {department ? department : (
            <button type="button" className="mh-add" onClick={() => open('place')}>Add department</button>
          )}
        </span>
        {(location || department) && (
          <button type="button" className="mh-pencil" aria-label={OPENABLE.place.ask}
            title={OPENABLE.place.ask} onClick={() => open('place')}>
            <Pencil size={14} />
          </button>
        )}
      </div>

      {group && (
        <form className="mh-form" onSubmit={submit}>
          {group.fields.map(({ key, label, placeholder }) => (
            <label className="ad-field" key={key}>
              <span>{label}</span>
              <input list={`${id}-${key}`} value={valueOf(key)} onChange={set(key)}
                placeholder={current(key) ? '' : placeholder || `No ${label.toLowerCase()} yet`} />
              <datalist id={`${id}-${key}`}>
                {(options[key] ?? []).map((o) => <option key={o} value={o} />)}
              </datalist>
            </label>
          ))}
          <span className="mh-form-actions">
            <Button type="submit" size="sm" loading={busy} disabled={!Object.keys(edits).length}>Save</Button>
            <Button type="button" variant="secondary" size="sm" disabled={busy} onClick={close}>Cancel</Button>
          </span>
        </form>
      )}
      {note && !group && <p className="mh-say" role="status">{note}</p>}
    </>
  );
}

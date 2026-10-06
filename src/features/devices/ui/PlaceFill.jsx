import { useId, useState } from 'react';
import Button from '../../../components/ui/Button';

const FIELDS = [
  { key: 'location', label: 'Location', placeholder: 'e.g. F1' },
  { key: 'department', label: 'Department', placeholder: '' },
];

const textOf = (value) => String(value ?? '').trim();

/**
 * A machine's location and department, set or corrected in place. A blank
 * field opens as a box; a set one shows its value and a Change button. Saving
 * is a correction of the record, written as a manual edit so the next import
 * leaves it alone -- a machine actually going to somebody else still goes
 * through Change owner, which is what the owner history records.
 */
export default function PlaceFill({ device, locations, departments, onSave, busy }) {
  const id = useId();
  const [opened, setOpened] = useState([]);
  const [values, setValues] = useState({});
  // What was just saved, shown until the register's re-read catches up: the
  // old value sitting there meanwhile looked like a save still running. The
  // page keys this component on the stored values, so the re-read resets it.
  const [saved, setSaved] = useState({});
  const [note, setNote] = useState('');

  const current = (key) => (key in saved ? saved[key] : textOf(device[key]));
  const editing = FIELDS.filter(({ key }) => !current(key) || opened.includes(key)).map(({ key }) => key);
  const valueOf = (key) => (key in values ? values[key] : current(key));

  const edits = Object.fromEntries(editing
    .map((key) => [key, textOf(valueOf(key))])
    // Locations are stored upper-case, so 'f2' over 'F2' is no change; a
    // department's capitalisation is the department's to correct.
    .filter(([key, value]) => value && (key === 'location'
      ? value.toUpperCase() !== current(key).toUpperCase()
      : value !== current(key))));

  const open = (key) => {
    setNote('');
    setOpened((list) => [...list, key]);
  };
  const close = () => {
    setOpened([]);
    setValues({});
  };
  const set = (key) => (event) => setValues((list) => ({ ...list, [key]: event.target.value }));

  const submit = async (event) => {
    event.preventDefault();
    if (!Object.keys(edits).length) return;
    if (await onSave(edits)) {
      const shown = { ...edits };
      if ('location' in shown) shown.location = shown.location.toUpperCase();
      setSaved((list) => ({ ...list, ...shown }));
      close();
      setNote('Saved.');
    }
  };

  const options = { location: locations, department: departments };

  return (
    <form className="dd-place" onSubmit={submit}>
      {FIELDS.map(({ key, label, placeholder }) => (editing.includes(key) ? (
        <label className="ad-field" key={key}>
          <span>{label}</span>
          <input list={`${id}-${key}`} value={valueOf(key)} onChange={set(key)}
            placeholder={current(key) ? '' : placeholder || `No ${key} yet`} />
          <datalist id={`${id}-${key}`}>{options[key].map((o) => <option key={o} value={o} />)}</datalist>
        </label>
      ) : (
        <span className="dd-place-value" key={key}>
          <span className="dd-place-label">{label}</span>
          <strong>{current(key)}</strong>
          <button type="button" className="dd-place-change" onClick={() => open(key)}>Change</button>
        </span>
      )))}
      {editing.length > 0 && (
        <span className="dd-place-actions">
          <Button type="submit" size="sm" loading={busy} disabled={!Object.keys(edits).length}>Save</Button>
          {opened.length > 0 && (
            <Button type="button" variant="secondary" size="sm" disabled={busy} onClick={close}>Cancel</Button>
          )}
        </span>
      )}
      {note && !editing.length && <span className="dd-place-say" role="status">{note}</span>}
    </form>
  );
}

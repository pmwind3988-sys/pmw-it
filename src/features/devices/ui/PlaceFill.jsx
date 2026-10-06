import { useId, useState } from 'react';
import Button from '../../../components/ui/Button';

/**
 * Fills in a location or department the machine has never had. Only the blank
 * ones are offered: filling a gap is a correction, while changing a place the
 * machine already has is a move, and moves go through Change owner so the
 * owner history records them.
 */
export default function PlaceFill({ device, locations, departments, onSave, busy }) {
  const id = useId();
  const [values, setValues] = useState({ location: '', department: '' });
  const missing = ['location', 'department'].filter((key) => !String(device[key] ?? '').trim());
  if (!missing.length) return null;

  const set = (key) => (event) => setValues((current) => ({ ...current, [key]: event.target.value }));
  const edits = Object.fromEntries(missing
    .map((key) => [key, values[key].trim()])
    .filter(([, value]) => value));

  const submit = (event) => {
    event.preventDefault();
    if (Object.keys(edits).length) onSave(edits);
  };

  return (
    <form className="dd-place" onSubmit={submit}>
      <span className="dd-place-say">
        No {missing.join(' or ')} recorded yet.
      </span>
      {missing.includes('location') && (
        <label className="ad-field">
          <span>Location</span>
          <input list={`${id}-locs`} value={values.location} onChange={set('location')} placeholder="e.g. F1" />
          <datalist id={`${id}-locs`}>{locations.map((l) => <option key={l} value={l} />)}</datalist>
        </label>
      )}
      {missing.includes('department') && (
        <label className="ad-field">
          <span>Department</span>
          <input list={`${id}-depts`} value={values.department} onChange={set('department')} />
          <datalist id={`${id}-depts`}>{departments.map((d) => <option key={d} value={d} />)}</datalist>
        </label>
      )}
      <Button type="submit" size="sm" loading={busy} disabled={!Object.keys(edits).length}>Save</Button>
    </form>
  );
}

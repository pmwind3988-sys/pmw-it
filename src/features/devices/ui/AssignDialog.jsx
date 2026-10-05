import { useEffect, useId, useRef, useState } from 'react';
import Button from '../../../components/ui/Button';
import { parseFormDate } from '../../forms/toChecklistItem';

const today = () => {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
};

/**
 * Who the machine goes to. Owner, location and department are suggested from
 * the register; anything else typed is taken as written. Escape closes it and
 * is stopped here, so it cannot also unwind the page behind.
 */
export default function AssignDialog({
  title, device, owners, locations, departments, onSubmit, onCancel, busy,
}) {
  const id = useId();
  const first = useRef(null);
  const [values, setValues] = useState({
    owner: '', location: device.location ?? '', department: device.department ?? '', on: today(), note: '',
  });
  const set = (key) => (event) => setValues((current) => ({ ...current, [key]: event.target.value }));

  useEffect(() => { first.current?.focus(); }, []);

  const submit = (event) => {
    event.preventDefault();
    onSubmit({ ...values, on: parseFormDate(values.on) ?? Date.now() });
  };

  return (
    <div className="ui-confirm" role="dialog" aria-modal="true" aria-labelledby={`${id}-t`}
      onKeyDown={(event) => { if (event.key === 'Escape') { event.stopPropagation(); onCancel(); } }}>
      <button type="button" className="ui-confirm-back" aria-label="Cancel" onClick={onCancel} />
      <form className="ui-confirm-box ad" onSubmit={submit}>
        <h2 id={`${id}-t`} className="ui-confirm-title">{title}</h2>
        {device.owner && <p className="ad-now">{device.owner} has it now.</p>}
        <label className="ad-field">
          <span>New owner</span>
          <input ref={first} required list={`${id}-owners`} value={values.owner} onChange={set('owner')} />
          <datalist id={`${id}-owners`}>{owners.map((o) => <option key={o} value={o} />)}</datalist>
        </label>
        <div className="ad-row">
          <label className="ad-field">
            <span>Location</span>
            <input list={`${id}-locs`} value={values.location} onChange={set('location')} />
            <datalist id={`${id}-locs`}>{locations.map((l) => <option key={l} value={l} />)}</datalist>
          </label>
          <label className="ad-field">
            <span>Department</span>
            <input list={`${id}-depts`} value={values.department} onChange={set('department')} />
            <datalist id={`${id}-depts`}>{departments.map((d) => <option key={d} value={d} />)}</datalist>
          </label>
          <label className="ad-field">
            <span>From</span>
            <input type="date" value={values.on} onChange={set('on')} />
          </label>
        </div>
        <label className="ad-field">
          <span>Note <em>(optional)</em></span>
          <textarea rows={2} value={values.note} onChange={set('note')} />
        </label>
        <p className="ad-summary">
          {device.owner ? `${device.owner}'s time with this machine ends` : 'It goes'} on the date above
          {values.owner ? `, and ${values.owner}'s begins.` : '.'} The specs and their history stay with the machine.
        </p>
        <div className="ui-confirm-actions">
          <Button variant="secondary" type="button" onClick={onCancel}>Cancel</Button>
          <Button type="submit" loading={busy} disabled={!values.owner.trim()}>{title.startsWith('Bring') ? 'Bring back' : 'Change owner'}</Button>
        </div>
      </form>
    </div>
  );
}

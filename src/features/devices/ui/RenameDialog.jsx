import { useEffect, useId, useRef, useState } from 'react';
import Button from '../../../components/ui/Button';
import { Pencil } from '../../../components/ui/Icons';
import { planRename } from '../sharepoint/updateDevice';

/**
 * A new computer name for one machine, in the app's own dialog. The reason a
 * name is refused shows as it is typed, and the old name is said to be kept,
 * because "where did the old one go" is the first thing anybody asks.
 */
export default function RenameDialog({
  device, devices, busy, error, onSave, onClose,
}) {
  const id = useId();
  const inputRef = useRef(null);
  const [name, setName] = useState(device.computerName ?? '');
  const plan = planRename(device, name, devices);
  const typed = name.trim() !== String(device.computerName ?? '').trim();

  useEffect(() => {
    inputRef.current?.select();
  }, []);

  useEffect(() => {
    const onKey = (event) => {
      if (event.key === 'Escape' && !busy) onClose();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [busy, onClose]);

  const submit = (event) => {
    event.preventDefault();
    if (!plan.error && !busy) onSave(plan.name);
  };

  return (
    <div className="ui-confirm" role="dialog" aria-modal="true" aria-labelledby={`${id}-title`}>
      <button type="button" className="ui-confirm-back" aria-label="Cancel" onClick={() => !busy && onClose()} />
      <form className="ui-confirm-box rn-box" onSubmit={submit}>
        <h2 className="ui-confirm-title" id={`${id}-title`}>
          <Pencil size={16} /> Rename {device.computerName}
        </h2>
        <p className="ui-confirm-body">
          The old name is kept under Previous names and in this machine&apos;s History as
          “Renamed”, and still finds it in search. The next scan report will not change it back.
          {!device.serialNumber && ' This machine has no serial on record, so a report that still '
            + 'carries the old name will arrive as a new machine.'}
        </p>
        <label className="rn-field" htmlFor={`${id}-name`}>
          New computer name
          <input id={`${id}-name`} ref={inputRef} className="rn-input" value={name}
            onChange={(event) => setName(event.target.value)} autoComplete="off" spellCheck={false}
            aria-invalid={typed && Boolean(plan.error)} aria-describedby={`${id}-why`} />
        </label>
        <p className="rn-why" id={`${id}-why`} role="status">
          {error || (typed && plan.error) || ''}
        </p>
        <div className="ui-confirm-actions">
          <Button variant="secondary" type="button" onClick={onClose} disabled={busy}>Cancel</Button>
          <Button type="submit" loading={busy} disabled={Boolean(plan.error)}>Rename</Button>
        </div>
      </form>
    </div>
  );
}

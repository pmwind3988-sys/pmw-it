import { useId, useState } from 'react';
import { NEW_MACHINE } from '../lifecycle/matchIncoming';

const nameKey = (name) => String(name ?? '').trim().toLowerCase();

/**
 * "Saves as" for one report in the import review: update the machine the
 * import matched, save it as a new machine, or update another saved machine
 * picked by name. `onChoose(undefined)` hands the row back to the automatic
 * match.
 */
export default function ImportTarget({ match, machines, onChoose }) {
  const id = useId();
  const [picking, setPicking] = useState(false);
  const [typed, setTyped] = useState('');

  const auto = match.auto ?? null;
  const pickedOther = match.chosen && match.existing && match.existing.id !== auto?.id;
  let value = 'auto';
  if (match.chosen && !match.existing) value = 'new';
  else if (pickedOther) value = `id:${match.existing.id}`;
  if (picking) value = 'other';

  const choose = (next) => {
    if (next === 'other') {
      setPicking(true);
      setTyped('');
      return;
    }
    setPicking(false);
    if (next === 'auto') onChoose(undefined);
    else if (next === 'new') onChoose(NEW_MACHINE);
  };

  const pick = (text) => {
    setTyped(text);
    const key = nameKey(text);
    if (!key) return;
    const found = machines.find((row) => nameKey(row.computerName) === key)
      ?? machines.find((row) => (row.previousNames ?? []).some((name) => nameKey(name) === key));
    if (found?.id !== null && found?.id !== undefined) {
      setPicking(false);
      onChoose(found.id);
    }
  };

  const label = `Saves ${match.device.computerName} as`;

  return (
    <div className="it-target">
      <select value={value} aria-label={label} onChange={(event) => choose(event.target.value)}>
        <option value="auto">{auto ? `Update ${auto.computerName}` : 'New machine'}</option>
        {auto && <option value="new">New machine</option>}
        {pickedOther && <option value={value}>Update {match.existing.computerName}</option>}
        <option value="other">{auto ? 'Update another machine…' : 'Update a saved machine…'}</option>
      </select>
      {picking && (
        <>
          <input
            className="it-target-find"
            list={`${id}-machines`}
            value={typed}
            placeholder="Type its computer name"
            aria-label={`Machine for ${match.device.computerName} to update`}
            autoFocus
            onChange={(event) => pick(event.target.value)}
            onKeyDown={(event) => { if (event.key === 'Escape') setPicking(false); }}
          />
          <datalist id={`${id}-machines`}>
            {machines.map((row) => <option key={row.id} value={row.computerName} />)}
          </datalist>
        </>
      )}
    </div>
  );
}

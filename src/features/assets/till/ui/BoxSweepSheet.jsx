import { useCallback, useMemo, useState } from 'react';
import TillSheet from './TillSheet';
import Button from '../../../../components/ui/Button';
import { Check } from '../../../../components/ui/Icons';
import { useTextScanner, SCAN_STATE } from '../../scan/useTextScanner';
import {
  newSweep, recordSweep, sweepSuggestions, sweepPending,
} from '../boxSweep';

const LABELS = {
  category: 'What it is', make: 'Make', model: 'Model', colour: 'Colour',
  connection: 'Connection', connector: 'Connector', length: 'Length', capacity: 'Capacity',
  screen: 'Screen', resolution: 'Resolution', power: 'Power',
};

/**
 * Sweep a box: keep the camera on it and turn it slowly. Every side is read,
 * and what the printing says comes up as suggestions — each one read at least
 * twice, each one a guess that can be switched off. Nothing is written until
 * "Use these".
 */
export default function BoxSweepSheet({ onCancel, onUse }) {
  const [sweep, setSweep] = useState(newSweep);
  const [off, setOff] = useState(() => new Set());

  const onLines = useCallback((lines) => setSweep((current) => recordSweep(current, lines)), []);
  const { videoRef, state } = useTextScanner({ active: true, maxPasses: Infinity, onLines });

  const found = useMemo(() => sweepSuggestions(sweep), [sweep]);
  const pending = useMemo(() => sweepPending(sweep), [sweep]);

  // One row per suggestion, in the order a person reads a box.
  const rows = [
    ...['category', 'make', 'model', 'colour'].filter((kind) => found[kind]).map((kind) => ({ id: kind, kind, value: found[kind] })),
    ...found.details.map((value) => ({ id: `detail:${value}`, kind: 'detail', value })),
  ];
  const confirmed = new Set(rows.map((row) => row.value));
  const waiting = pending.filter((entry) => !confirmed.has(entry.value)).slice(0, 4);

  const toggle = (id) => setOff((current) => {
    const next = new Set(current);
    if (next.has(id)) next.delete(id); else next.add(id);
    return next;
  });

  const use = () => {
    const keep = (id) => !off.has(id);
    onUse({
      category: keep('category') ? found.category : '',
      make: keep('make') ? found.make : '',
      model: keep('model') ? found.model : '',
      colour: keep('colour') ? found.colour : '',
      details: found.details.filter((value) => keep(`detail:${value}`)),
    });
  };

  const broken = state === SCAN_STATE.DENIED || state === SCAN_STATE.UNAVAILABLE || state === SCAN_STATE.NO_READER;
  const chosen = rows.filter((row) => !off.has(row.id)).length;

  return (
    <TillSheet title="Sweep the box" onClose={onCancel} wide>
      <p className="till-sheet-lede">
        {broken
          ? 'This browser cannot read text from the camera here. Type the details instead.'
          : 'Turn the box slowly so the camera sees every side. Each value has to be read twice before it is offered.'}
      </p>

      <div className="till-scan">
        <div className="till-viewfinder as-viewfinder">
          <video ref={videoRef} playsInline muted className="as-video" />
          {state === SCAN_STATE.STARTING && <p className="as-camera-msg">Starting the camera and the reader…</p>}
        </div>
        <span className="till-sweep-sides">Sides read: {sweep.passes}</span>
      </div>

      <ul className="till-sweep-list">
        {rows.map((row) => (
          <li key={row.id}>
            <label className="till-check">
              <input type="checkbox" checked={!off.has(row.id)} onChange={() => toggle(row.id)} />
              <span className="till-sweep-kind">{row.kind === 'detail' ? 'Detail' : LABELS[row.kind]}</span>
              <strong>{row.value}</strong>
            </label>
            <span className="till-sweep-tag">guessed · read twice</span>
          </li>
        ))}
        {waiting.map((entry) => (
          <li key={`${entry.kind}:${entry.value}`} className="till-sweep-waiting">
            <span className="till-sweep-kind">{LABELS[entry.kind] ?? 'Detail'}</span>
            <span>{entry.value}</span>
            <span className="till-sweep-tag">reading…</span>
          </li>
        ))}
        {!rows.length && !waiting.length && <li className="till-sweep-waiting">Nothing read yet.</li>}
      </ul>

      <Button icon={Check} className="till-cta" disabled={!chosen} onClick={use}>
        {chosen ? `Use these (${chosen})` : 'Keep turning the box'}
      </Button>
    </TillSheet>
  );
}

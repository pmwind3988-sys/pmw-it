import { useCallback, useRef, useState } from 'react';
import TillSheet from './TillSheet';
import TillCamera from './TillCamera';
import Button from '../../../../components/ui/Button';
import { Check } from '../../../../components/ui/Icons';
import { signalAccepted, signalDuplicate } from '../../scan/feedback';
import { createReadGate, guessKind } from '../readGate';
import {
  newRun, addToRun, withoutSerial, undoLast, runSize, RUN_RESULT,
} from '../serialRun';

/**
 * Ten of the same thing, each with its own serial: scan them one after
 * another and the count is how many were scanned. The camera here is told it
 * wants SERIALS, so the shop barcode on each box is passed over without
 * asking. A thing with its sticker gone is one tap, and still counts.
 */
export default function SerialRunSheet({ title, assets, drafts, onCancel, onDone }) {
  const [run, setRun] = useState(newRun);
  const [typed, setTyped] = useState('');
  const [flash, setFlash] = useState(null);
  const [choices, setChoices] = useState([]);
  const gateRef = useRef(createReadGate());
  const runRef = useRef(run);
  const timer = useRef(null);

  const say = useCallback((kind, text) => {
    if (kind === 'ok') signalAccepted(); else signalDuplicate();
    setFlash({ kind, text });
    clearTimeout(timer.current);
    timer.current = setTimeout(() => setFlash(null), 1600);
  }, []);

  const take = useCallback((raw) => {
    const out = addToRun(runRef.current, raw, { assets, drafts });
    if (out.result === RUN_RESULT.ADDED) {
      runRef.current = out.run;
      setRun(out.run);
      say('ok', `${runSize(out.run)} · ${out.run.serials[out.run.serials.length - 1]}`);
    } else if (out.result === RUN_RESULT.REPEAT) say('ask', 'Already scanned in this run');
    else if (out.result === RUN_RESULT.ON_RECEIPT) say('bad', 'That serial is already on the receipt');
    else if (out.result === RUN_RESULT.REGISTERED) say('bad', `Already registered: ${out.asset.title || 'that serial'}`);
  }, [assets, drafts, say]);

  const onCodes = useCallback((codes, meta = {}) => {
    const { accept, choices: unsure } = gateRef.current(codes, { aimed: Boolean(meta.aimed), expect: 'serial' });
    for (const code of accept) take(code);
    if (unsure.length) setChoices(unsure);
  }, [take]);

  const change = (next) => { runRef.current = next; setRun(next); };
  const count = runSize(run);

  return (
    <TillSheet title={title} onClose={onCancel} wide>
      <p className="till-sheet-lede">Scan each one’s serial. The count is how many you scan.</p>
      <div className="till-scan">
        <TillCamera
          active
          onCodes={onCodes}
          flash={flash}
          choices={choices.map((code) => ({ code, kind: guessKind(code) }))}
          onPick={(code) => { setChoices([]); take(code); }}
          onDismiss={() => setChoices([])}
        />
        <form
          className="till-entry"
          onSubmit={(event) => { event.preventDefault(); if (typed.trim()) take(typed); setTyped(''); }}
        >
          <label htmlFor="run-code" className="sr-only">Type a serial</label>
          <input id="run-code" value={typed} onChange={(event) => setTyped(event.target.value)} placeholder="…or type a serial" autoComplete="off" autoCapitalize="characters" spellCheck={false} />
          <button type="submit" className="till-add">Add</button>
        </form>
      </div>

      <div className="till-run-bar">
        <strong className="till-mono">{count}</strong>
        <span>{run.serials.length} with a serial{run.without ? ` · ${run.without} without` : ''}</span>
        <span className="till-run-actions">
          <button type="button" className="till-chip" onClick={() => change(withoutSerial(run))}>One without a serial</button>
          <button type="button" className="till-chip" onClick={() => change(undoLast(run))} disabled={!count}>Undo last</button>
        </span>
      </div>

      {run.serials.length > 0 && (
        <ol className="till-run-list till-mono" reversed>
          {[...run.serials].reverse().slice(0, 8).map((serial) => <li key={serial}>{serial}</li>)}
          {run.serials.length > 8 && <li className="till-run-more">and {run.serials.length - 8} more</li>}
        </ol>
      )}

      <Button icon={Check} className="till-cta" disabled={!count} onClick={() => onDone(run)}>
        {count ? `Add ${count}` : 'Scan the first one'}
      </Button>
    </TillSheet>
  );
}

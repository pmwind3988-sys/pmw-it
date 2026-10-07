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
 * wants SERIALS, and a shop barcode (or one this line is already known by) is
 * set aside as the model's, never counted as an item. A thing with its sticker
 * gone is one tap, and still counts.
 *
 * `existing` are the serials the line already holds, listed first so the
 * person can see what is already there; `boxCodes` are the barcodes the line
 * is known by.
 */
export default function SerialRunSheet({
  title, assets, drafts, existing = [], boxCodes = [], onCancel, onDone,
}) {
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
    timer.current = setTimeout(() => setFlash(null), 1800);
  }, []);

  const take = useCallback((raw) => {
    const out = addToRun(runRef.current, raw, { assets, drafts, boxCodes });
    runRef.current = out.run;
    setRun(out.run);
    if (out.result === RUN_RESULT.ADDED) {
      const serial = out.run.serials[out.run.serials.length - 1];
      say('ok', `Item ${existing.length + runSize(out.run)} · ${serial}`);
    } else if (out.result === RUN_RESULT.BOX_CODE) say('ask', 'Shop barcode — kept for the model, not a serial');
    else if (out.result === RUN_RESULT.REPEAT) say('ask', 'Already scanned in this run');
    else if (out.result === RUN_RESULT.ON_RECEIPT) say('bad', 'That serial is already on the receipt');
    else if (out.result === RUN_RESULT.REGISTERED) say('bad', `Already registered: ${out.asset.title || 'that serial'}`);
  }, [assets, drafts, boxCodes, existing.length, say]);

  const onCodes = useCallback((codes, meta = {}) => {
    const { accept, choices: unsure } = gateRef.current(codes, { aimed: Boolean(meta.aimed), expect: 'serial' });
    for (const code of accept) take(code);
    if (unsure.length) setChoices(unsure);
  }, [take]);

  const change = (next) => { runRef.current = next; setRun(next); };
  const count = runSize(run);

  // Everything on the line, item by item: what was already there, then this
  // run's serials and the ones counted without, newest at the top.
  const listed = [
    ...existing.map((serial, i) => ({ id: `old:${serial}`, n: i + 1, serial, old: true })),
    ...run.serials.map((serial, i) => ({ id: `new:${serial}`, n: existing.length + i + 1, serial, old: false })),
    ...Array.from({ length: run.without }, (_, i) => ({ id: `none:${i}`, n: existing.length + run.serials.length + i + 1, serial: '', old: false })),
  ].reverse();

  return (
    <TillSheet title={title} onClose={onCancel} wide>
      <p className="till-sheet-lede">Scan each one’s serial. The count is how many you scan. Shop barcodes are skipped.</p>
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
        <span>
          new{run.serials.length ? ` · ${run.serials.length} with a serial` : ''}{run.without ? ` · ${run.without} without` : ''}
          {existing.length ? ` · ${existing.length} already on this line` : ''}
        </span>
        <span className="till-run-actions">
          <button type="button" className="till-chip" onClick={() => change(withoutSerial(run))}>One without a serial</button>
          <button type="button" className="till-chip" onClick={() => change(undoLast(run))} disabled={!count}>Undo last</button>
        </span>
      </div>

      {listed.length > 0 && (
        <ul className="till-run-list" aria-label="Serials on this line">
          {listed.map((item) => (
            <li key={item.id} className={item.old ? 'till-run-old' : 'till-run-new'}>
              <span className="till-run-n till-mono">{item.n}</span>
              <span className="till-mono">{item.serial || 'no serial'}</span>
              <span className="till-run-where">{item.old ? 'already on this line' : 'this run'}</span>
            </li>
          ))}
        </ul>
      )}

      {run.boxCodes.length > 0 && (
        <p className="till-run-box">Model barcode kept: <span className="till-mono">{run.boxCodes.join(', ')}</span></p>
      )}

      <Button icon={Check} className="till-cta" disabled={!count} onClick={() => onDone(run)}>
        {count ? `Add ${count}` : 'Scan the first one'}
      </Button>
    </TillSheet>
  );
}

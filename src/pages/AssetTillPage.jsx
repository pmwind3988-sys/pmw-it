import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Link, useSearchParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { ErrorBanner } from '../components/ui/Surfaces';
import {
  Camera, Check, ScanLine, Truck, UserPlus, Inbox, Boxes,
} from '../components/ui/Icons';
import { useAssets, SHAREPOINT_SITE_URL } from '../features/assets/useAssets';
import { useHandovers } from '../features/assets/useHandovers';
import { useSharePointToken } from '../hooks/useRequests';
import { signalAccepted, signalDuplicate, signalDone } from '../features/assets/scan/feedback';
import { applyScannedFields } from '../features/assets/scan/textScan';
import TextScanSheet from '../features/assets/ui/TextScanSheet';
import PersonPicker from '../features/assets/ui/PersonPicker';
import { newBatch, addDraft, replaceDraft, removeDraft } from '../features/assets/draft/batch';
import { newDraft } from '../features/assets/draft/draftAsset';
import { saveBatch, loadPhoto } from '../features/assets/store/assetDb';
import { readPref, writePref } from '../features/assets/till/prefs';
import { newCount, addCounted, COUNT_RESULT } from '../features/assets/till/count';
import { useStoredBatch } from '../features/assets/till/ui/useStoredBatch';
import CountPanel from '../features/assets/till/ui/CountPanel';
import { saveBatchToSharePoint, remainingDrafts } from '../features/assets/sharepoint/saveBatch';
import { commitHandover, commitReturn } from '../features/assets/sharepoint/writeHandover';
import { newBasket, removeLine, setQuantity, isUnitLine } from '../features/assets/handover/basket';
import { isOverdue, isOpen } from '../features/assets/handover/availability';
import { categoriesIn } from '../features/assets/categories';
import { assetTitle } from '../features/assets/identity';
import { TRACKED, CONDITIONS } from '../features/assets/assetKinds';
import {
  scanIn, addModel, setKind, setLineField, needsKind, needsSerial, holdsFor, nextTag, matchRegister,
  noCodeDraft, itemCount as itemsIn, STOCK_RESULT,
} from '../features/assets/till/stockIn';
import {
  scanOut, addAsset, refusalsFor, sendable, itemCount as itemsOut, withTerms, DUE_CHOICES, OUT_RESULT,
} from '../features/assets/till/handOut';
import {
  scanBack, addHandover, chooseHolder, setReturnQuantity, toReturns, itemCount as itemsBack, BACK_RESULT,
} from '../features/assets/till/takeBack';
import { tillSearch, quickKeys } from '../features/assets/till/tillSearch';
import { createReadGate, guessKind } from '../features/assets/till/readGate';
import { findScanTarget } from '../features/assets/handover/scanMatch';
import TillCamera from '../features/assets/till/ui/TillCamera';
import TillReceipt from '../features/assets/till/ui/TillReceipt';
import TillSheet from '../features/assets/till/ui/TillSheet';
import CantScanSheet from '../features/assets/till/ui/CantScanSheet';
import NoCodeSheet from '../features/assets/till/ui/NoCodeSheet';
import SignStep from '../features/assets/till/ui/SignStep';
import DoneReceipt from '../features/assets/till/ui/DoneReceipt';

/**
 * The asset till: stock in, hand out and take back on one screen, the way a
 * cashier works — scan things onto a receipt, check it, one button to finish.
 *
 * Every decision about what a scan MEANS is in `features/assets/till/` and
 * tested there; this page holds the state and says what happened.
 */

const MODES = [
  { id: 'in', label: 'Stock in', icon: Truck },
  { id: 'out', label: 'Hand out', icon: UserPlus },
  { id: 'back', label: 'Take back', icon: Inbox },
  { id: 'count', label: 'Count', icon: Boxes },
];

const PHASE = {
  provisioning: 'Setting up SharePoint',
  reading: 'Checking the register',
  photos: 'Uploading photos',
  writing: 'Saving',
  logging: 'Recording changes',
  delivery: 'Recording the delivery',
  signature: 'Saving the signature',
  updating: 'Updating the register',
};

const RECEIPT_TITLE = {
  in: 'Receipt · stock in', out: 'Receipt · handover', back: 'Receipt · return', count: 'Receipt · catch-up count',
};
const RECEIPT_EMPTY = {
  in: 'Scan the boxes as they come off the trolley. The same model again counts up.',
  out: 'Scan or type what they are taking. Anything that cannot go out says why straight away.',
  back: 'Scan whatever came back. The till finds who had it.',
  count: 'Count what is already here, room by room. Things with no serial are fine — tick it and take a photo.',
};

const COMMON_KINDS = ['Laptop', 'Monitor', 'Mouse', 'Keyboard', 'Cable', 'Accessory'];
const BATCH_KEY = 'tillBatchId';
const COUNT_KEY = 'tillCountId';
const CAMERA_KEY = 'tillCamera';


/** Phones scan with the camera; a desk has a scanner gun and keeps it off. */
function cameraByDefault() {
  const saved = readPref(CAMERA_KEY);
  if (saved != null) return saved === 'on';
  return typeof window !== 'undefined' && window.matchMedia?.('(pointer: coarse)').matches;
}

const BADGES = {
  Laptop: 'LT', Desktop: 'DT', Monitor: 'MN', Printer: 'PR', 'Docking Station': 'DK', Phone: 'PH',
  Tab: 'TB', Network: 'NW', 'PC Part': 'PC', Keyboard: 'KB', Mouse: 'MS', Cable: 'CB', Adapter: 'AD',
  Accessory: 'AC', Other: 'OT',
};
/** Enter in a line's own box saves it, the way leaving the box does. */
const commitOnEnter = (event) => { if (event.key === 'Enter') { event.preventDefault(); event.currentTarget.blur(); } };
const badgeOf = (category) => BADGES[category] ?? String(category || '?').slice(0, 2).toUpperCase();
const plural = (n, word) => `${n} ${word}${n === 1 ? '' : 's'}`;
const firstName = (name = '') => name.split(' ')[0] || name;
const stamp = () => new Date().toLocaleString('en-MY', {
  weekday: 'short', day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit',
});

export default function AssetTillPage() {
  const [params, setParams] = useSearchParams();
  const mode = MODES.some((entry) => entry.id === params.get('mode')) ? params.get('mode') : 'in';

  const { instance } = useMsal();
  const getToken = useSharePointToken();
  const { assets, error: assetsError, reload: reloadAssets } = useAssets();
  const { handovers, reload: reloadHandovers } = useHandovers();
  const assetsById = useMemo(() => new Map(assets.map((asset) => [asset.id, asset])), [assets]);

  // The receipts. Each mode keeps its own, so switching tab to check
  // something does not throw away a half-scanned delivery. A delivery and a
  // count are batches kept on this device; the other two need the network
  // anyway.
  const [batch, setBatch, finishBatch] = useStoredBatch(BATCH_KEY, newBatch);
  const [countBatch, setCountBatch, finishCount] = useStoredBatch(COUNT_KEY, newCount);
  const [basket, setBasket] = useState(() => newBasket());
  const [missesOut, setMissesOut] = useState([]);
  const [terms, setTerms] = useState({ loan: false, dueChoice: '2w' });
  const [backLines, setBackLines] = useState([]);
  const [missesBack, setMissesBack] = useState([]);
  const [condition, setCondition] = useState('Good');

  const [typed, setTyped] = useState('');
  const [flash, setFlash] = useState(null);
  const [nudge, setNudge] = useState(false);
  const [cameraOn, setCameraOn] = useState(cameraByDefault);
  const [sheet, setSheet] = useState(null);
  const [headerOpen, setHeaderOpen] = useState(false);
  const [signature, setSignature] = useState(null);
  const [progress, setProgress] = useState(null);
  const [failure, setFailure] = useState('');
  const [done, setDone] = useState(null);

  const inputRef = useRef(null);
  const flashTimer = useRef(null);
  const labelDraftRef = useRef(null);
  const gateRef = useRef(createReadGate());
  // Codes the camera could not choose between, offered to the person.
  const [choices, setChoices] = useState([]);

  // What the scan handlers read. A camera frame can carry several codes, and
  // each must see the receipt the previous one left — not the one this render
  // closed over.
  const live = useRef({});
  useEffect(() => {
    live.current = { mode, batch, basket, backLines, assets, handovers };
  });

  useEffect(() => () => clearTimeout(flashTimer.current), []);

  const say = useCallback((kind, text) => {
    if (kind === 'ok') signalAccepted(); else signalDuplicate();
    setFlash({ kind, text });
    clearTimeout(flashTimer.current);
    flashTimer.current = setTimeout(() => setFlash(null), 1800);
  }, []);

  // ── What one code does, per mode ─────────────────────────────────────────

  const handleCode = useCallback((raw, format = '') => {
    const state = live.current;

    if (state.mode === 'in') {
      const out = scanIn(state.batch, raw, state.assets, format);
      live.current = { ...state, batch: out.batch };
      setBatch(out.batch);
      const name = out.draft ? assetTitle(out.draft) : '';
      if (out.result === STOCK_RESULT.ADDED) say('ok', `Added · ${name}`);
      else if (out.result === STOCK_RESULT.COUNTED) say('ok', `+1 · ${name} (now ${out.draft.quantity})`);
      else if (out.result === STOCK_RESULT.FILLED) say('ok', `Serial added · ${name}`);
      else if (out.result === STOCK_RESULT.NEW) say('ask', 'New code — say what it is');
      else if (out.result === STOCK_RESULT.DUPLICATE) say('ask', 'Already on this receipt');
      else if (out.result === STOCK_RESULT.KNOWN) say('ask', `${out.asset.title || 'That'} is already in the register`);
      return;
    }

    if (state.mode === 'out') {
      const out = scanOut(state.basket, raw, state.assets);
      if (out.result === OUT_RESULT.MISSING) {
        setMissesOut((list) => (list.some((miss) => miss.code === raw) ? list : [...list, { id: `${raw}-${Date.now()}`, code: String(raw).trim() }]));
        say('bad', 'Not in the register');
        return;
      }
      if (out.result === OUT_RESULT.DUPLICATE) { say('ask', 'Already on this receipt'); return; }
      live.current = { ...state, basket: out.basket };
      setBasket(out.basket);
      const reason = refusalsFor(out.basket, state.assets).get(out.line?.lineId);
      if (reason) say('bad', reason);
      else say('ok', out.result === OUT_RESULT.COUNTED ? `+1 · ${out.line.itemTitle} (now ${out.line.quantity})` : `Added · ${out.line.itemTitle}`);
      return;
    }

    const out = scanBack(state.backLines, raw, state.assets, state.handovers);
    if (out.result === BACK_RESULT.MISSING || out.result === BACK_RESULT.NOT_OUT) {
      const reason = out.result === BACK_RESULT.MISSING ? 'Not in the register' : 'Not out with anyone — it is in stock';
      setMissesBack((list) => [...list, { id: `${raw}-${Date.now()}`, code: String(raw).trim(), name: out.asset?.title, reason }]);
      say('bad', reason);
      return;
    }
    if (out.result === BACK_RESULT.DUPLICATE) { say('ask', 'Already on this receipt'); return; }
    if (out.result === BACK_RESULT.ALL_BACK) { say('bad', `${out.line.from} has no more of these out`); return; }
    live.current = { ...state, backLines: out.lines };
    setBackLines(out.lines);
    if (out.result === BACK_RESULT.CHOOSE) say('ask', 'Several people have these — whose is it?');
    else say('ok', `${out.result === BACK_RESULT.COUNTED ? '+1 · ' : ''}Back from ${out.line.from}`);
  }, [say, setBatch]);

  /**
   * Every camera read goes through the gate (`till/readGate.js`): read twice,
   * aim wins, and when it cannot tell which barcode is meant it asks. While
   * the last delivery line is waiting for its serial, a serial-shaped code is
   * taken without asking.
   */
  const onCodes = useCallback((codes, meta = {}) => {
    const state = live.current;
    const last = state.batch?.drafts[state.batch.drafts.length - 1];
    const expect = state.mode === 'in' && last && needsSerial(last) && !needsKind(last) ? 'serial' : null;
    const { accept, choices: unsure } = gateRef.current(codes, { aimed: Boolean(meta.aimed), expect });
    for (const code of accept) handleCode(code);
    if (unsure.length) {
      signalDuplicate();
      setChoices(unsure);
    }
  }, [handleCode]);

  const pickChoice = (code) => {
    setChoices([]);
    handleCode(code);
  };

  /** What a code is, in words, so the right one can be picked. */
  const describe = (code) => {
    if (mode === 'in') {
      const hit = matchRegister(assets, code);
      if (hit) return hit.by === 'part' ? `Model number of ${hit.asset.title || hit.asset.model}` : `Already registered: ${hit.asset.title}`;
    } else {
      const hit = findScanTarget(assets, code);
      if (hit) return `In the register: ${hit.asset.title}`;
    }
    return guessKind(code);
  };

  const onQuiet = useCallback((quiet) => { if (quiet) setNudge(true); }, []);

  const submitTyped = (event) => {
    event.preventDefault();
    const code = typed.trim();
    if (!code) return;
    setTyped('');
    handleCode(code);
    inputRef.current?.focus();
  };

  // ── Picks without a scan: search results, quick keys, the holder list ────

  const results = useMemo(
    () => tillSearch(mode, typed, { assets, handovers }),
    [mode, typed, assets, handovers],
  );

  const pick = (result) => {
    setTyped('');
    if (result.kind === 'model') {
      const out = addModel(batch, result.asset);
      setBatch(out.batch);
      say('ok', out.result === STOCK_RESULT.COUNTED ? `+1 · ${result.name}` : `Added · ${result.name}`);
    } else if (result.kind === 'asset') {
      const out = addAsset(basket, result.asset);
      if (out.result === OUT_RESULT.DUPLICATE) { say('ask', 'Already on this receipt'); return; }
      setBasket(out.basket);
      say('ok', `Added · ${result.name}`);
    } else {
      const out = addHandover(backLines, result.handover, assets);
      if (out.result === BACK_RESULT.DUPLICATE) { say('ask', 'Already on this receipt'); return; }
      setBackLines(out.lines);
      say('ok', `Back from ${out.line.from}`);
    }
    inputRef.current?.focus();
  };

  /** One counted thing onto the count. Returns what happened, so the panel
   *  knows whether to clear its serial and photo for the next one. */
  const countOne = (entry) => {
    const out = addCounted(countBatch, entry, assets);
    const name = out.draft ? assetTitle(out.draft) : '';
    if (out.result === COUNT_RESULT.ADDED) {
      setCountBatch(out.batch);
      say('ok', `Added · ${name}`);
    } else if (out.result === COUNT_RESULT.COUNTED) {
      setCountBatch(out.batch);
      say('ok', `+${entry.quantity} · ${name} (now ${out.draft.quantity})`);
    } else if (out.result === COUNT_RESULT.DUPLICATE) {
      say('ask', 'That serial is already on this count');
    } else if (out.result === COUNT_RESULT.KNOWN) {
      say('ask', `${out.asset.title || 'That serial'} is already in the register`);
    }
    return out.result;
  };

  const keys = useMemo(() => quickKeys(assets, handovers), [assets, handovers]);
  const outNow = useMemo(() => handovers
    .filter((row) => isOpen(row))
    .sort((a, b) => Number(isOverdue(b)) - Number(isOverdue(a)))
    .slice(0, 8), [handovers]);

  const pressKey = (asset) => {
    if (mode === 'in') pick({ kind: 'model', asset, name: asset.title || asset.model });
    else pick({ kind: 'asset', asset, name: asset.title || asset.model });
  };

  // ── Reading the label instead ─────────────────────────────────────────────

  const takeLabel = (values, guessed, extras) => {
    if (mode !== 'in') {
      const code = values.serialNumber || values.assetTag || values.partNumber;
      if (code) {
        setSheet(null);
        handleCode(code);
      }
      return;
    }
    // A label carries several values and the sheet hands them over one at a
    // time; they all belong to the ONE thing being held up.
    setBatch((current) => {
      const existing = current.drafts.find((draft) => draft.localId === labelDraftRef.current);
      const base = existing ?? newDraft({ scanSource: 'Camera' });
      const { record } = applyScannedFields(base, values, guessed, extras, { byHand: true });
      labelDraftRef.current = record.localId;
      return existing ? replaceDraft(current, record) : addDraft(current, record);
    });
    signalAccepted();
  };

  const closeLabel = () => {
    labelDraftRef.current = null;
    setSheet(null);
  };

  // ── Checkout ──────────────────────────────────────────────────────────────

  const holds = useMemo(() => holdsFor(batch, assets), [batch, assets]);
  const countHolds = useMemo(() => holdsFor(countBatch, assets), [countBatch, assets]);
  const refusals = useMemo(() => refusalsFor(basket, assets), [basket, assets]);
  const busy = Boolean(progress);
  const busyLabel = progress ? `${PHASE[progress.phase] ?? 'Working'}…` : '';
  const who = () => {
    const account = instance.getActiveAccount();
    return account?.username ?? account?.name ?? '';
  };

  const run = async (work) => {
    setFailure('');
    setProgress({ phase: 'reading' });
    try {
      await work();
    } catch (thrown) {
      setFailure(thrown.message || 'That did not go through. Nothing on the receipt was lost — try again.');
    } finally {
      setProgress(null);
    }
  };

  /** A delivery or a count: both are batches, and save the same way. */
  const saveDrafts = (kind) => run(async () => {
    const counting = kind === 'count';
    const target = counting ? countBatch : batch;
    const setTarget = counting ? setCountBatch : setBatch;
    await saveBatch(target).catch(() => {});
    const token = (await getToken()).accessToken;
    const report = await saveBatchToSharePoint({
      siteUrl: SHAREPOINT_SITE_URL, token, batch: target, photoFor: loadPhoto, savedBy: who(), onProgress: setProgress,
    });
    reloadAssets();

    const left = remainingDrafts(target, report);
    const savedRows = target.drafts.filter((draft) => !left.some((keep) => keep.localId === draft.localId));
    if (!left.length) {
      if (counting) await finishCount(target, newCount());
      else await finishBatch(target, newBatch({ purchase: { supplier: target.purchase.supplier } }));
    } else {
      setTarget({ ...target, drafts: left });
    }
    const places = [...new Set(savedRows.map((draft) => draft.location).filter(Boolean))];
    signalDone();
    setDone({
      mode: kind,
      title: left.length ? `Saved ${plural(savedRows.length, 'line')}` : 'Saved to the register',
      when: stamp(),
      meta: counting
        ? [
          { k: 'Counted in', v: places.join(', ') || '—' },
          { k: 'Without a serial', v: String(savedRows.filter((draft) => draft.noSerial).length) },
        ]
        : [
          { k: 'Supplier', v: target.purchase.supplier || '—' },
          { k: 'DO number', v: target.purchase.doNumber || 'To follow' },
        ],
      rows: savedRows.map((draft) => ({ id: draft.localId, name: assetTitle(draft), sub: draft.serialNumber || draft.assetTag || '', qty: draft.trackingMode === TRACKED ? 1 : draft.quantity })),
      total: savedRows.reduce((sum, draft) => sum + (draft.trackingMode === TRACKED ? 1 : draft.quantity), 0),
      warning: left.length ? `${plural(left.length, 'line')} could not be saved and are still on the receipt.` : '',
      next: left.length ? 'Back to the receipt' : (counting ? 'Keep counting' : 'Next delivery'),
      link: left.length ? { to: `/assets/batch/${target.id}`, label: 'Open the full review' } : { to: '/assets', label: 'Open the register' },
    });
  });

  const handOver = () => run(async () => {
    const token = (await getToken()).accessToken;
    const toSend = sendable(withTerms(basket, terms), refusals);
    const report = await commitHandover({
      siteUrl: SHAREPOINT_SITE_URL, token, basket: toSend, issuedBy: who(), signature, onProgress: setProgress,
    });
    reloadAssets();
    reloadHandovers();

    const blocked = new Set(report.blocked.map((entry) => entry.line.lineId));
    const sent = toSend.lines.filter((line) => !blocked.has(line.lineId));
    const due = withTerms(basket, terms).dueOn;
    signalDone();
    setSheet(null);
    setDone({
      mode: 'out',
      title: `Handed to ${basket.person.name}`,
      when: stamp(),
      meta: [
        { k: 'To', v: basket.person.email || basket.person.name },
        { k: 'Terms', v: terms.loan ? (due ? `Loan · back by ${new Date(due).toLocaleDateString('en-MY', { day: 'numeric', month: 'short', year: 'numeric' })}` : 'Loan · no date') : 'Theirs to keep' },
        { k: 'Signature', v: report.signed ? 'Signed' : (report.signatureFailed ? 'Did not upload' : 'Not signed') },
      ],
      rows: sent.map((line) => ({
        id: line.lineId,
        name: line.itemTitle,
        sub: line.serialNumber || line.unitLabel || (line.trackingMode === TRACKED ? assetsById.get(line.assetId)?.serialNumber : '') || '',
        qty: line.quantity,
      })),
      total: sent.reduce((sum, line) => sum + line.quantity, 0),
      warning: blocked.size ? `${plural(blocked.size, 'line')} could not go out and are still on the receipt.` : '',
      next: 'Next person',
      link: { to: `/assets/people/${encodeURIComponent(basket.person.email)}`, label: `Open ${firstName(basket.person.name)}’s page` },
    });
    if (!report.signatureFailed) setSignature(null);
    // Only what the write itself refused stays, to be pressed again. Lines
    // refused before checkout were left off on purpose and mean nothing to
    // the next person at the till.
    setBasket((current) => ({ ...current, lines: current.lines.filter((line) => blocked.has(line.lineId)) }));
    setMissesOut([]);
  });

  const recordReturn = () => run(async () => {
    const token = (await getToken()).accessToken;
    const returns = toReturns(backLines, condition);
    const report = await commitReturn({
      siteUrl: SHAREPOINT_SITE_URL, token, returns, returnedBy: who(), signature: null, onProgress: setProgress,
    });
    reloadAssets();
    reloadHandovers();

    const blocked = new Set(report.blocked.map((entry) => entry.entry.handoverId));
    const sent = backLines.filter((line) => line.handoverId != null && !blocked.has(line.handoverId));
    signalDone();
    setDone({
      mode: 'back',
      title: 'Return recorded',
      when: stamp(),
      meta: [
        { k: 'From', v: [...new Set(sent.map((line) => line.from))].join(', ') || '—' },
        { k: 'Condition', v: condition },
      ],
      rows: sent.map((line) => ({ id: line.lineId, name: line.name, sub: `From ${line.from}`, qty: line.quantity })),
      total: sent.reduce((sum, line) => sum + line.quantity, 0),
      warning: blocked.size ? `${plural(blocked.size, 'line')} could not be recorded and are still on the receipt.` : '',
      next: 'Next return',
      link: { to: '/assets/people', label: 'Open who has what' },
    });
    setBackLines((current) => current.filter((line) => line.handoverId == null || blocked.has(line.handoverId)));
    setMissesBack([]);
  });

  const nextJob = () => {
    if (done?.mode === 'out' && !basket.lines.length) {
      setBasket(newBasket());
      setMissesOut([]);
    }
    setDone(null);
    inputRef.current?.focus();
  };

  // ── Receipt rows, per mode ────────────────────────────────────────────────

  const categories = useMemo(() => categoriesIn(assets), [assets]);

  const draftRows = (theBatch, setTheBatch, theHolds) => theBatch.drafts.map((draft) => {
    const unnamed = needsKind(draft);
    const hold = theHolds.get(draft.localId);
    const update = (field) => (event) => setTheBatch((current) => setLineField(current, draft.localId, field, event.target.value));
    const bulk = draft.trackingMode !== TRACKED;
    return {
      id: draft.localId,
      badge: unnamed ? '?' : badgeOf(draft.category),
      name: unnamed ? 'New code' : assetTitle(draft),
      sub: [
        draft.serialNumber ? `S/N ${draft.serialNumber}` : (draft.assetTag ? `Label ${draft.assetTag}` : (draft.partNumber ? `Part ${draft.partNumber}` : '')),
        draft.noSerial && (draft.photoId ? 'No serial · photo taken' : 'No serial · no photo'),
        draft.location,
      ].filter(Boolean).join(' · '),
      note: hold,
      tone: hold ? 'ask' : 'ok',
      qty: bulk ? draft.quantity : 1,
      onInc: bulk && !unnamed ? () => setTheBatch((current) => setLineField(current, draft.localId, 'quantity', draft.quantity + 1)) : null,
      onDec: () => setTheBatch((current) => setLineField(current, draft.localId, 'quantity', Math.max(1, draft.quantity - 1))),
      onRemove: () => setTheBatch((current) => removeDraft(current, draft.localId)),
      extra: (unnamed || hold) && (
        <div className="till-fix">
          {unnamed && (
            <div className="till-chips">
              {COMMON_KINDS.map((kind) => (
                <button key={kind} type="button" className="till-chip" onClick={() => setTheBatch((current) => setKind(current, draft.localId, kind))}>{kind}</button>
              ))}
              <select
                className="till-chip"
                value=""
                aria-label="Another category"
                onChange={(event) => event.target.value && setTheBatch((current) => setKind(current, draft.localId, event.target.value))}
              >
                <option value="">More…</option>
                {categories.filter((name) => !COMMON_KINDS.includes(name)).map((name) => <option key={name}>{name}</option>)}
              </select>
            </div>
          )}
          {!unnamed && needsSerial(draft) && (
            <input className="till-inline" placeholder="Serial number" aria-label="Serial number" onBlur={update('serialNumber')} onKeyDown={commitOnEnter} defaultValue="" />
          )}
          {!unnamed && !needsSerial(draft) && !String(draft.model ?? '').trim() && (
            <input className="till-inline" placeholder="Make and model" aria-label="Model" onBlur={update('model')} onKeyDown={commitOnEnter} defaultValue="" />
          )}
        </div>
      ),
    };
  });
  const rowsIn = draftRows(batch, setBatch, holds);
  const rowsCount = draftRows(countBatch, setCountBatch, countHolds);

  const rowsOut = [
    ...basket.lines.map((line) => {
      const reason = refusals.get(line.lineId);
      const countable = line.trackingMode !== TRACKED && !isUnitLine(line);
      const asset = assetsById.get(line.assetId);
      const serial = line.serialNumber || (line.trackingMode === TRACKED ? asset?.serialNumber : '');
      const label = line.trackingMode === TRACKED ? asset?.assetTag : '';
      return {
        id: line.lineId,
        badge: badgeOf(line.category),
        name: line.itemTitle,
        sub: [serial && `S/N ${serial}`, label, !serial && (line.unitLabel || line.category)].filter(Boolean).join(' · '),
        note: reason,
        tone: reason ? 'bad' : 'ok',
        qty: line.quantity,
        onInc: countable ? () => setBasket((current) => setQuantity(current, line.lineId, line.quantity + 1)) : null,
        onDec: () => setBasket((current) => setQuantity(current, line.lineId, Math.max(1, line.quantity - 1))),
        onRemove: () => setBasket((current) => removeLine(current, line.lineId)),
      };
    }),
    ...missesOut.map((miss) => ({
      id: miss.id, badge: '?', name: 'Unknown code', sub: miss.code, tone: 'bad',
      note: 'Not in the register — stock it in first', qty: null,
      onRemove: () => setMissesOut((list) => list.filter((entry) => entry.id !== miss.id)),
    })),
  ];

  const rowsBack = [
    ...backLines.map((line) => ({
      id: line.lineId,
      badge: line.handoverId ? '↩' : '?',
      name: line.name,
      sub: line.handoverId ? `From ${line.from}` : 'Several people have these',
      note: line.handoverId ? '' : 'Whose is it?',
      tone: line.handoverId ? 'ok' : 'ask',
      qty: line.handoverId ? line.quantity : null,
      onInc: line.handoverId && !line.single ? () => setBackLines((current) => setReturnQuantity(current, line.lineId, line.quantity + 1)) : null,
      onDec: () => setBackLines((current) => setReturnQuantity(current, line.lineId, line.quantity - 1)),
      onRemove: () => setBackLines((current) => current.filter((entry) => entry.lineId !== line.lineId)),
      extra: !line.handoverId && (
        <div className="till-chips">
          {line.choices.map((choice) => (
            <button
              key={choice.handoverId}
              type="button"
              className="till-chip"
              onClick={() => setBackLines((current) => chooseHolder(current, line.lineId, choice.handoverId, handovers, assets))}
            >
              {choice.from} · {choice.outstanding}
            </button>
          ))}
        </div>
      ),
    })),
    ...missesBack.map((miss) => ({
      id: miss.id, badge: '?', name: miss.name || 'Unknown code', sub: miss.code, tone: 'bad', note: miss.reason, qty: null,
      onRemove: () => setMissesBack((list) => list.filter((entry) => entry.id !== miss.id)),
    })),
  ];

  // ── Footer, per mode ──────────────────────────────────────────────────────

  let rows; let total; let note = ''; let noteBad = false; let cta; let ctaOk; let onCheckout;
  if (mode === 'in') {
    rows = rowsIn;
    total = itemsIn(batch);
    const unanswered = holds.size;
    cta = !batch.drafts.length ? 'Scan something to start'
      : unanswered ? `${plural(unanswered, 'line')} ${unanswered === 1 ? 'needs' : 'need'} an answer`
        : `Save ${plural(total, 'item')} to the register`;
    ctaOk = batch.drafts.length > 0 && !unanswered;
    note = 'Kept on this device until you save';
    onCheckout = () => saveDrafts('in');
  } else if (mode === 'count') {
    rows = rowsCount;
    total = itemsIn(countBatch);
    const unanswered = countHolds.size;
    cta = !countBatch.drafts.length ? 'Count something to start'
      : unanswered ? `${plural(unanswered, 'line')} ${unanswered === 1 ? 'needs' : 'need'} an answer`
        : `Save ${plural(total, 'item')} to the register`;
    ctaOk = countBatch.drafts.length > 0 && !unanswered;
    note = 'Kept on this device until you save — save as often as you like';
    onCheckout = () => saveDrafts('count');
  } else if (mode === 'out') {
    rows = rowsOut;
    total = itemsOut(basket, refusals);
    const left = refusals.size + missesOut.length;
    if (left) { note = `${plural(left, 'line')} can’t go out and will be left off`; noteBad = true; }
    cta = !total ? 'Scan what they are getting'
      : !basket.person ? 'Choose who is getting it'
        : `Hand ${total} to ${firstName(basket.person.name)} · sign`;
    ctaOk = total > 0;
    onCheckout = () => setSheet(basket.person ? 'sign' : 'person');
  } else {
    rows = rowsBack;
    total = itemsBack(backLines);
    const asking = backLines.filter((line) => !line.handoverId).length;
    if (asking) { note = `Say whose ${asking === 1 ? 'it is' : 'they are'} on ${plural(asking, 'line')}`; noteBad = true; }
    cta = total ? `Record return of ${total}` : 'Scan what came back';
    ctaOk = total > 0;
    onCheckout = recordReturn;
  }

  const switchMode = (id) => {
    setParams((current) => { const next = new URLSearchParams(current); next.set('mode', id); return next; }, { replace: true });
    setTyped('');
    setFlash(null);
    setChoices([]);
    setDone(null);
  };

  const toggleCamera = () => {
    setCameraOn((on) => { writePref(CAMERA_KEY, on ? 'off' : 'on'); return !on; });
  };

  const counts = {
    in: itemsIn(batch), out: itemsOut(basket, refusals), back: itemsBack(backLines), count: itemsIn(countBatch),
  };
  const purchase = batch.purchase;
  const setPurchase = (field, value) => setBatch((current) => ({ ...current, purchase: { ...current.purchase, [field]: value } }));

  return (
    <AppShell title="Till" subtitle="Scan things onto the receipt. One button to finish.">
      <div className="till">
        {assetsError && <ErrorBanner message={assetsError} onRetry={reloadAssets} />}
        {failure && <ErrorBanner message={failure} onRetry={onCheckout} />}

        <div className="till-modes" role="tablist" aria-label="What are you doing?">
          {MODES.map(({ id, label, icon: Icon }) => (
            <button
              key={id}
              type="button"
              role="tab"
              aria-selected={mode === id}
              className={mode === id ? 'till-mode on' : 'till-mode'}
              onClick={() => switchMode(id)}
            >
              <Icon size={16} />
              {label}
              {counts[id] > 0 && <span className="till-count till-mono">{counts[id]}</span>}
            </button>
          ))}
        </div>

        <div className="till-grid">
          <div className="till-left">
            {mode === 'count' ? (
              <>
                <CountPanel assets={assets} drafts={countBatch.drafts} categories={categories} onAdd={countOne} />
                {flash && <div className={`till-flash till-flash-inline till-flash-${flash.kind}`} role="status">{flash.text}</div>}
              </>
            ) : (
            <>
            <div className="till-scan">
              {cameraOn && !done && (
                <TillCamera
                  active={!sheet}
                  onCodes={onCodes}
                  flash={flash}
                  onQuiet={onQuiet}
                  choices={choices.map((code) => ({ code, kind: describe(code) }))}
                  onPick={pickChoice}
                  onDismiss={() => setChoices([])}
                />
              )}

              <form className="till-entry" onSubmit={submitTyped}>
                <label htmlFor="till-code" className="sr-only">Scan or type a code</label>
                <input
                  id="till-code"
                  ref={inputRef}
                  value={typed}
                  onChange={(event) => setTyped(event.target.value)}
                  placeholder={mode === 'back' ? 'Scan, or type a code or name' : 'Scan, or type a code'}
                  autoComplete="off"
                  autoCapitalize="characters"
                  spellCheck={false}
                  // A scanner gun types into whatever has focus; on a desk this
                  // box should be it the moment the page opens.
                  autoFocus={!cameraOn}
                />
                <button type="submit" className="till-add">Add</button>
                <button
                  type="button"
                  className={nudge ? 'till-cant on' : 'till-cant'}
                  onClick={() => { setNudge(false); setSheet('help'); }}
                >
                  Can’t scan?
                </button>
              </form>

              {!cameraOn && flash && (
                <div className={`till-flash till-flash-inline till-flash-${flash.kind}`} role="status">{flash.text}</div>
              )}

              {results.length > 0 && (
                <ul className="till-results" aria-label="Matches">
                  {results.map((result) => (
                    <li key={result.id}>
                      <button type="button" onClick={() => pick(result)}>
                        <strong>{result.name}</strong>
                        <span className="till-mono">{result.sub}</span>
                      </button>
                    </li>
                  ))}
                  <li className="till-results-foot">Enter uses exactly what you typed</li>
                </ul>
              )}

              <div className="till-scan-foot">
                <button type="button" className="till-link" onClick={toggleCamera}>
                  <Camera size={14} /> {cameraOn ? 'Turn the camera off' : 'Use the camera'}
                </button>
              </div>
            </div>

            {mode !== 'back' && keys.length > 0 && (
              <section className="till-keys" aria-label="Quick keys">
                <h2 className="till-kicker">Quick keys</h2>
                <div className="till-keys-grid">
                  {keys.map(({ asset, left }) => {
                    const empty = mode === 'out' && left < 1;
                    return (
                      <button key={asset.id} type="button" className={empty ? 'till-key empty' : 'till-key'} disabled={empty} onClick={() => pressKey(asset)}>
                        <span className="till-mono">{badgeOf(asset.category)}</span>
                        <strong>{asset.title || asset.model}</strong>
                        <span>{mode === 'out' ? (left ? `${left} in stock` : 'Out of stock') : 'Add one'}</span>
                      </button>
                    );
                  })}
                </div>
              </section>
            )}

            {mode === 'back' && outNow.length > 0 && (
              <section className="till-keys" aria-label="Out right now">
                <h2 className="till-kicker">Out right now</h2>
                <ul className="till-outnow">
                  {outNow.map((row) => (
                    <li key={row.id}>
                      <span>
                        <strong>{row.personName || row.personEmail}</strong>
                        <span>{row.itemTitle}{isOverdue(row) ? ' · overdue' : ''}</span>
                      </span>
                      <button type="button" className="till-chip" onClick={() => pick({ kind: 'handover', handover: row })}>Take back</button>
                    </li>
                  ))}
                </ul>
              </section>
            )}
            </>
            )}
          </div>

          <div className="till-right">
            {done ? (
              <DoneReceipt done={done} onNext={nextJob} />
            ) : (
              <>
                {mode === 'in' && (
                  <section className="till-context">
                    <button type="button" className="till-context-row" aria-expanded={headerOpen} onClick={() => setHeaderOpen((open) => !open)}>
                      <span className="till-kicker">Delivery</span>
                      <span className="till-context-sum">
                        {purchase.supplier || 'No supplier yet'} · {purchase.doNumber ? `DO ${purchase.doNumber}` : 'DO to follow'}
                      </span>
                      <span className="till-link">{headerOpen ? 'Done' : 'Edit'}</span>
                    </button>
                    {headerOpen && (
                      <div className="till-context-form">
                        <label className="till-field till-span"><span>Supplier</span>
                          <input value={purchase.supplier ?? ''} onChange={(event) => setPurchase('supplier', event.target.value)} />
                        </label>
                        <label className="till-field"><span>DO number</span>
                          <input value={purchase.doNumber ?? ''} onChange={(event) => setPurchase('doNumber', event.target.value)} placeholder="Later is fine" />
                        </label>
                        <label className="till-field"><span>PO number</span>
                          <input value={purchase.poNumber ?? ''} onChange={(event) => setPurchase('poNumber', event.target.value)} placeholder="Later is fine" />
                        </label>
                        <label className="till-check till-span">
                          <input type="checkbox" checked={Boolean(purchase.detailsPending)} onChange={(event) => setPurchase('detailsPending', event.target.checked)} />
                          The paperwork has not arrived yet — details to follow
                        </label>
                      </div>
                    )}
                  </section>
                )}

                {mode === 'out' && (
                  <section className="till-context">
                    {basket.person ? (
                      <PersonPicker person={basket.person} onChange={(person) => setBasket((current) => ({ ...current, person }))} />
                    ) : (
                      <button type="button" className="till-who" onClick={() => setSheet('person')}>
                        <UserPlus size={16} /> Who is getting it?
                      </button>
                    )}
                    <div className="till-chips">
                      <button type="button" className={terms.loan ? 'till-chip' : 'till-chip on'} aria-pressed={!terms.loan} onClick={() => setTerms((t) => ({ ...t, loan: false }))}>Keep</button>
                      <button type="button" className={terms.loan ? 'till-chip on' : 'till-chip'} aria-pressed={terms.loan} onClick={() => setTerms((t) => ({ ...t, loan: true }))}>Loan</button>
                      {terms.loan && <span className="till-chip-gap" aria-hidden="true" />}
                      {terms.loan && DUE_CHOICES.map((choice) => (
                        <button
                          key={choice.id}
                          type="button"
                          className={terms.dueChoice === choice.id ? 'till-chip on' : 'till-chip'}
                          aria-pressed={terms.dueChoice === choice.id}
                          onClick={() => setTerms((t) => ({ ...t, dueChoice: choice.id }))}
                        >
                          {choice.label}
                        </button>
                      ))}
                    </div>
                  </section>
                )}

                {mode === 'back' && (
                  <section className="till-context">
                    <span className="till-kicker">Coming back in</span>
                    <div className="till-chips">
                      {CONDITIONS.filter((name) => name !== 'New' && name !== 'Retired').map((name) => (
                        <button key={name} type="button" className={condition === name ? 'till-chip on' : 'till-chip'} aria-pressed={condition === name} onClick={() => setCondition(name)}>{name}</button>
                      ))}
                    </div>
                  </section>
                )}

                <TillReceipt
                  title={RECEIPT_TITLE[mode]}
                  rows={rows}
                  empty={RECEIPT_EMPTY[mode]}
                />

                <footer className="till-foot">
                  <div className="till-total">
                    <span className="till-kicker">{mode === 'back' ? 'Coming back' : 'Items'}</span>
                    <strong className="till-mono">{total}</strong>
                    {note && <span className={noteBad ? 'till-foot-note bad' : 'till-foot-note'}>{note}</span>}
                  </div>
                  <Button
                    icon={mode === 'out' ? Check : ScanLine}
                    className="till-cta"
                    disabled={!ctaOk}
                    loading={busy && sheet !== 'sign'}
                    onClick={onCheckout}
                  >
                    {busy && sheet !== 'sign' ? busyLabel : cta}
                  </Button>
                  {mode === 'in' && holds.size > 0 && (
                    <Link className="till-link" to={`/assets/batch/${batch.id}`} onClick={() => saveBatch(batch).catch(() => {})}>
                      Or finish it on the full review page
                    </Link>
                  )}
                </footer>
              </>
            )}
          </div>
        </div>
      </div>

      {sheet === 'help' && (
        <CantScanSheet
          mode={mode}
          onClose={() => setSheet(null)}
          onReadLabel={() => setSheet('label')}
          onType={() => { setSheet(null); setTimeout(() => inputRef.current?.focus(), 0); }}
          onNoCode={() => setSheet('nocode')}
        />
      )}

      {sheet === 'label' && (
        <TextScanSheet title="Read the label or screen" onCancel={closeLabel} onUse={takeLabel} />
      )}

      {sheet === 'nocode' && (
        <NoCodeSheet
          categories={categories}
          tag={nextTag(assets, batch.drafts)}
          onCancel={() => setSheet(null)}
          onAdd={(entry) => {
            const draft = noCodeDraft(entry);
            setBatch((current) => addDraft(current, draft));
            setSheet(null);
            say('ok', `Added · ${assetTitle(draft)}`);
          }}
        />
      )}

      {sheet === 'person' && (
        <TillSheet title="Who is getting it?" onClose={() => setSheet(null)}>
          <PersonPicker
            person={null}
            onChange={(person) => {
              setBasket((current) => ({ ...current, person }));
              setSheet(null);
            }}
          />
        </TillSheet>
      )}

      {sheet === 'sign' && basket.person && (
        <SignStep
          person={basket.person}
          count={total}
          terms={terms.loan ? `On loan · ${DUE_CHOICES.find((c) => c.id === terms.dueChoice)?.label ?? ''}` : 'Theirs to keep'}
          rows={sendable(basket, refusals).lines.map((line) => ({
            id: line.lineId,
            name: line.itemTitle,
            sub: line.serialNumber || (line.trackingMode === TRACKED ? assetsById.get(line.assetId)?.serialNumber : ''),
            qty: line.quantity,
          }))}
          signature={signature}
          onSignature={setSignature}
          onConfirm={handOver}
          onClose={() => !busy && setSheet(null)}
          busy={busy}
          busyLabel={busyLabel}
          failure={failure}
        />
      )}
    </AppShell>
  );
}

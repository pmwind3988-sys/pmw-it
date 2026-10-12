import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Link, useNavigate, useSearchParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { ErrorBanner } from '../components/ui/Surfaces';
import {
  Camera, Check, ScanLine, Truck, UserPlus, Inbox, Boxes, ClipboardList,
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
import SerialRunSheet from '../features/assets/till/ui/SerialRunSheet';
import SameAsSearch from '../features/assets/till/ui/SameAsSearch';
import BoxSweepSheet from '../features/assets/till/ui/BoxSweepSheet';
import ChoiceSheet from '../features/assets/till/ui/ChoiceSheet';
import { checklistFromTill } from '../features/assets/till/checklistFromTill';
import { useOrgDirectory } from '../hooks/useOrgDirectory';
import { snapshotOptions } from '../features/forms/formOptions';
import { IN, OUT, INDIVIDUAL, modeLabel } from '../features/forms/checklistForm';
import { createLink, draftLink, linkUrl } from '../features/forms/sharepoint/checklistLinks';
import { saveBatchToSharePoint, remainingDrafts } from '../features/assets/sharepoint/saveBatch';
import { commitHandover, commitReturn } from '../features/assets/sharepoint/writeHandover';
import { newBasket, removeLine, setQuantity, isUnitLine } from '../features/assets/handover/basket';
import { isOverdue, isOpen } from '../features/assets/handover/availability';
import { categoriesIn } from '../features/assets/categories';
import { assetTitle } from '../features/assets/identity';
import { parseUnits } from '../features/assets/units';
import { TRACKED, CONDITIONS } from '../features/assets/assetKinds';
import {
  scanIn, addModel, setKind, setLineField, needsKind, needsSerial, holdsFor, nextTag, matchRegister,
  noCodeDraft, itemCount as itemsIn, STOCK_RESULT, linkToModel, addSerialsTo, serialCount, applySweep,
} from '../features/assets/till/stockIn';
import {
  scanOut, addAsset, refusalsFor, sendable, itemCount as itemsOut, withTerms, DUE_CHOICES, OUT_RESULT,
} from '../features/assets/till/handOut';
import {
  scanBack, addHandover, chooseHolder, setReturnQuantity, toReturns, itemCount as itemsBack, BACK_RESULT,
} from '../features/assets/till/takeBack';
import { tillSearch, quickKeys } from '../features/assets/till/tillSearch';
import {
  inStash, findInStash, searchStash, findOutOfStash, searchOutOfStash, machineRow, machineTitle, machineSub,
} from '../features/assets/till/stashMachines';
import { useDevices } from '../features/devices/useDevices';
import { performLifecycle } from '../features/devices/sharepoint/writeLifecycle';
import { ACTIONS } from '../features/devices/lifecycle/planLifecycle';
import { SPARE, statusOf, inFleet } from '../features/devices/lifecycle/status';
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

// The checklists the till sent for a signature, signed or still waiting.
const TILL_CHECKLISTS = '/asset-checklist/links?from=till';

const PHASE = {
  provisioning: 'Setting up SharePoint',
  reading: 'Checking the register',
  photos: 'Uploading photos',
  writing: 'Saving',
  logging: 'Recording changes',
  delivery: 'Recording the delivery',
  signature: 'Saving the signature',
  link: 'Creating the checklist link',
  machine: 'Updating the device list',
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
  const navigate = useNavigate();
  const mode = MODES.some((entry) => entry.id === params.get('mode')) ? params.get('mode') : 'in';

  const { instance } = useMsal();
  const getToken = useSharePointToken();
  const { assets, error: assetsError, reload: reloadAssets } = useAssets();
  const { handovers, reload: reloadHandovers } = useHandovers();
  // Spare machines in the IT Stash can be handed out too (`till/stashMachines.js`).
  const { devices, reload: reloadDevices } = useDevices();
  const assetsById = useMemo(() => new Map(assets.map((asset) => [asset.id, asset])), [assets]);
  // Entity and department choices a checklist link carries for the employee.
  const directory = useOrgDirectory();
  const linkOptions = useMemo(() => (directory ? snapshotOptions(directory) : null), [directory]);

  // The receipts. Each mode keeps its own, so switching tab to check
  // something does not throw away a half-scanned delivery. A delivery and a
  // count are batches kept on this device; the other two need the network
  // anyway.
  const [batch, setBatch, finishBatch] = useStoredBatch(BATCH_KEY, newBatch);
  const [countBatch, setCountBatch, finishCount] = useStoredBatch(COUNT_KEY, newCount);
  const [basket, setBasket] = useState(() => newBasket());
  const [missesOut, setMissesOut] = useState([]);
  // Machines from the IT Stash on the handover receipt, beside the basket.
  const [machines, setMachines] = useState([]);
  const [terms, setTerms] = useState({ loan: false, dueChoice: '2w' });
  const [backLines, setBackLines] = useState([]);
  const [missesBack, setMissesBack] = useState([]);
  // Machines somebody had, coming back to the IT Stash.
  const [machinesBack, setMachinesBack] = useState([]);

  // `?machine=<id>`: opened from a machine's own page, which puts it on the
  // receipt as soon as the device list has loaded. Adjusted during render,
  // once, rather than from an effect.
  const [wantedMachine, setWantedMachine] = useState(() => params.get('machine'));
  if (wantedMachine && devices.length) {
    setWantedMachine(null);
    const device = devices.find((entry) => String(entry.id) === wantedMachine);
    if (device && mode === 'out' && statusOf(device) === SPARE) setMachines((list) => (list.some((entry) => entry.id === device.id) ? list : [...list, device]));
    if (device && mode === 'back' && inFleet(device)) setMachinesBack((list) => (list.some((entry) => entry.id === device.id) ? list : [...list, device]));
  }
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
  // A serial run onto one counted line: { localId, which: 'in' | 'count', title }.
  const [runFor, setRunFor] = useState(null);
  // A box sweep onto one line: { localId, which }.
  const [sweepFor, setSweepFor] = useState(null);

  // What the scan handlers read. A camera frame can carry several codes, and
  // each must see the receipt the previous one left — not the one this render
  // closed over.
  const live = useRef({});
  useEffect(() => {
    live.current = { mode, batch, basket, backLines, assets, handovers, devices, machines, machinesBack };
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
        // Not in the asset register, but it may be a machine in the IT Stash.
        const machine = findInStash(state.devices, raw);
        if (machine) {
          if (state.machines.some((device) => device.id === machine.id)) { say('ask', 'Already on this receipt'); return; }
          live.current = { ...state, machines: [...state.machines, machine] };
          setMachines((list) => [...list, machine]);
          say('ok', `Added from the IT Stash · ${machineTitle(machine)}`);
          return;
        }
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
    if (out.result === BACK_RESULT.MISSING) {
      // Not in the asset register, but it may be a machine somebody has.
      const machine = findOutOfStash(state.devices, raw);
      if (machine) {
        if (state.machinesBack.some((device) => device.id === machine.id)) { say('ask', 'Already on this receipt'); return; }
        live.current = { ...state, machinesBack: [...state.machinesBack, machine] };
        setMachinesBack((list) => [...list, machine]);
        say('ok', `Back to the IT Stash · ${machineTitle(machine)}`);
        return;
      }
    }
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

  const results = useMemo(() => {
    const found = tillSearch(mode, typed, { assets, handovers });
    if (mode === 'out') return [...found, ...searchStash(devices, typed, machines)];
    if (mode === 'back') return [...found, ...searchOutOfStash(devices, typed, machinesBack)];
    return found;
  }, [mode, typed, assets, handovers, devices, machines, machinesBack]);

  const addMachine = (device) => {
    if (machines.some((entry) => entry.id === device.id)) { say('ask', 'Already on this receipt'); return; }
    setMachines((list) => [...list, device]);
    say('ok', `Added from the IT Stash · ${machineTitle(device)}`);
  };
  const addMachineBack = (device) => {
    if (machinesBack.some((entry) => entry.id === device.id)) { say('ask', 'Already on this receipt'); return; }
    setMachinesBack((list) => [...list, device]);
    say('ok', `Back to the IT Stash · ${machineTitle(device)}`);
  };

  const pick = (result) => {
    setTyped('');
    if (result.kind === 'model') {
      const out = addModel(batch, result.asset);
      setBatch(out.batch);
      say('ok', out.result === STOCK_RESULT.COUNTED ? `+1 · ${result.name}` : `Added · ${result.name}`);
    } else if (result.kind === 'machine') {
      if (mode === 'back') addMachineBack(result.device); else addMachine(result.device);
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
      const added = (entry.serials?.length ?? 0) + (entry.without ?? 0) || entry.quantity;
      say('ok', `+${added} · ${name} (now ${out.draft.quantity})`);
    } else if (out.result === COUNT_RESULT.DUPLICATE) {
      say('ask', 'That serial is already on this count');
    } else if (out.result === COUNT_RESULT.KNOWN) {
      say('ask', `${out.asset.title || 'That serial'} is already in the register`);
    }
    return out.result;
  };

  const keys = useMemo(() => quickKeys(assets, handovers), [assets, handovers]);
  const stashLeft = useMemo(
    () => inStash(devices).filter((device) => !machines.some((entry) => entry.id === device.id)),
    [devices, machines],
  );
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

  /**
   * A checklist link for what was just recorded: pre-filled, locked to what
   * the till scanned, and tied to those handover rows so the employee's
   * signature lands on them. The handover is already recorded either way; a
   * link that cannot be made is reported, never a reason to undo it.
   */
  const makeChecklistLink = async (token, {
    formMode, person, rows, kind, ids, machines: onLink = [],
  }) => {
    setProgress({ phase: 'link' });
    try {
      const values = checklistFromTill({ formMode, person, rows });
      const link = {
        ...draftLink({ formMode, values, editable: ['position'], options: linkOptions, expiresInDays: 14 }),
        // The machines are kept on the link so it can say where each went
        // and open it; only `ids` get the employee's signature.
        handovers: {
          kind,
          ids,
          ...(onLink.length ? {
            machines: onLink.map((device) => ({ id: device.id, name: machineTitle(device), serial: device.serialNumber || '' })),
          } : {}),
        },
      };
      const account = instance.getActiveAccount();
      const created = await createLink({
        siteUrl: SHAREPOINT_SITE_URL, token, link,
        createdByName: account?.name ?? '', createdByEmail: account?.username ?? '',
      });
      return { url: linkUrl(window.location.origin, created.code), mode: modeLabel(formMode) };
    } catch (thrown) {
      return { error: thrown.message || 'The checklist link could not be created' };
    }
  };

  const serialOfLine = (line) => line.serialNumber || line.unitLabel
    || (line.trackingMode === TRACKED ? assetsById.get(line.assetId)?.serialNumber : '') || '';

  const handOver = (formMode = null) => run(async () => {
    const token = (await getToken()).accessToken;
    const toSend = sendable(withTerms(basket, terms), refusals);
    // Signed on the checklist link instead: no phone signature now — theirs
    // arrives on these rows when they sign. A receipt of stash machines only
    // has no handover rows to write.
    const report = toSend.lines.length
      ? await commitHandover({
        siteUrl: SHAREPOINT_SITE_URL, token, basket: toSend, issuedBy: who(), signature: formMode ? null : signature, onProgress: setProgress,
      })
      : { blocked: [], handoverIds: [] };
    if (toSend.lines.length) {
      reloadAssets();
      reloadHandovers();
    }

    // Stash machines go to the person in the device list — the same move as
    // "Assign" on the device page, refused if somebody took it meanwhile.
    const given = [];
    const stuck = [];
    for (const device of machines) {
      setProgress({ phase: 'machine' });
      try {
        await performLifecycle({
          siteUrl: SHAREPOINT_SITE_URL,
          token,
          deviceId: device.id,
          action: ACTIONS.ASSIGN,
          input: {
            owner: basket.person.name,
            location: device.location,
            note: `Handed out at the till${basket.person.email ? ` to ${basket.person.email}` : ''}`,
          },
          expectedStatus: SPARE,
          recordedBy: who(),
        });
        given.push(device);
      } catch (thrown) {
        stuck.push({ device, reason: thrown.message || 'could not be saved' });
      }
    }
    if (machines.length) reloadDevices();

    const blocked = new Set(report.blocked.map((entry) => entry.line.lineId));
    const sent = toSend.lines.filter((line) => !blocked.has(line.lineId));
    const due = withTerms(basket, terms).dueOn;
    const checklist = formMode && (sent.length || given.length)
      ? await makeChecklistLink(token, {
        formMode,
        person: basket.person,
        rows: [
          ...given.map(machineRow),
          ...sent.map((line) => ({
            category: line.category, name: line.itemTitle, serial: serialOfLine(line), qty: line.quantity,
          })),
        ],
        kind: 'issue',
        ids: report.handoverIds ?? [],
        machines: given,
      })
      : null;
    signalDone();
    setSheet(null);
    setDone({
      mode: 'out',
      checklist,
      title: `Handed to ${basket.person.name}`,
      when: stamp(),
      meta: [
        { k: 'To', v: basket.person.email || basket.person.name },
        { k: 'Terms', v: terms.loan ? (due ? `Loan · back by ${new Date(due).toLocaleDateString('en-MY', { day: 'numeric', month: 'short', year: 'numeric' })}` : 'Loan · no date') : 'Theirs to keep' },
        { k: 'Signature', v: checklist?.url ? 'Waiting for the checklist' : (report.signed ? 'Signed' : (report.signatureFailed ? 'Did not upload' : 'Not signed')) },
      ],
      rows: [
        ...given.map((device) => ({ id: `machine:${device.id}`, name: machineTitle(device), sub: machineSub(device), qty: 1 })),
        ...sent.map((line) => ({
          id: line.lineId,
          name: line.itemTitle,
          sub: serialOfLine(line),
          qty: line.quantity,
        })),
      ],
      total: given.length + sent.reduce((sum, line) => sum + line.quantity, 0),
      warning: [
        blocked.size ? `${plural(blocked.size, 'line')} could not go out and are still on the receipt.` : '',
        ...stuck.map(({ device, reason }) => `${machineTitle(device)} is still on the receipt: ${reason}`),
        checklist?.error ? `The handover is recorded, but the checklist link could not be made: ${checklist.error}` : '',
      ].filter(Boolean).join(' '),
      next: 'Next person',
      // A machine went to them in the DEVICE list, not the asset register:
      // with nothing else handed over, that is where to look for it.
      link: !sent.length && given.length === 1
        ? { to: `/devices/${given[0].id}`, label: `Open ${machineTitle(given[0])}` }
        : { to: `/assets/people/${encodeURIComponent(basket.person.email)}`, label: `Open ${firstName(basket.person.name)}’s page` },
    });
    if (!report.signatureFailed) setSignature(null);
    // Only what the write itself refused stays, to be pressed again. Lines
    // refused before checkout were left off on purpose and mean nothing to
    // the next person at the till.
    setBasket((current) => ({ ...current, lines: current.lines.filter((line) => blocked.has(line.lineId)) }));
    setMachines(stuck.map(({ device }) => device));
    setMissesOut([]);
  });

  const recordReturn = (withLink = false) => run(async () => {
    const token = (await getToken()).accessToken;
    const returns = toReturns(backLines, condition);
    const report = returns.length
      ? await commitReturn({
        siteUrl: SHAREPOINT_SITE_URL, token, returns, returnedBy: who(), signature: null, onProgress: setProgress,
      })
      : { blocked: [], returnedIds: [] };
    if (returns.length) {
      reloadAssets();
      reloadHandovers();
    }

    // Machines go back to the IT Stash in the device list -- the device
    // page's "To stash", which ends the owner's stint. Refused if somebody
    // moved it meanwhile.
    const stashed = [];
    const stuck = [];
    for (const device of machinesBack) {
      setProgress({ phase: 'machine' });
      try {
        await performLifecycle({
          siteUrl: SHAREPOINT_SITE_URL,
          token,
          deviceId: device.id,
          action: ACTIONS.TO_STASH,
          input: { note: `Taken back at the till · ${condition}` },
          expectedStatus: statusOf(device),
          recordedBy: who(),
        });
        stashed.push(device);
      } catch (thrown) {
        stuck.push({ device, reason: thrown.message || 'could not be saved' });
      }
    }
    if (machinesBack.length) reloadDevices();

    const blocked = new Set(report.blocked.map((entry) => entry.entry.handoverId));
    const sent = backLines.filter((line) => line.handoverId != null && !blocked.has(line.handoverId));
    const checklist = withLink && (sent.length || stashed.length)
      ? await makeChecklistLink(token, {
        formMode: OUT,
        person: { name: sent[0]?.from || stashed[0]?.owner || '', title: '' },
        rows: [...stashed.map(machineRow), ...sent.map((line) => {
          const row = handovers.find((entry) => entry.id === line.handoverId);
          const asset = assets.find((entry) => entry.assetKey === line.assetKey);
          return {
            category: asset?.category ?? row?.category ?? '',
            name: asset?.title || row?.itemTitle || line.name,
            serial: row?.serialNumber || (asset?.trackingMode === TRACKED ? asset.serialNumber : '') || '',
            qty: line.quantity,
          };
        })],
        kind: 'return',
        ids: report.returnedIds ?? [],
        machines: stashed,
      })
      : null;
    signalDone();
    setDone({
      mode: 'back',
      checklist,
      title: 'Return recorded',
      when: stamp(),
      meta: [
        { k: 'From', v: [...new Set([...stashed.map((device) => device.owner || 'Unknown owner'), ...sent.map((line) => line.from)])].join(', ') || '—' },
        { k: 'Condition', v: condition },
      ],
      rows: [
        ...stashed.map((device) => ({ id: `machine:${device.id}`, name: machineTitle(device), sub: `From ${device.owner || 'unknown owner'} · to the IT Stash`, qty: 1 })),
        ...sent.map((line) => ({ id: line.lineId, name: line.name, sub: `From ${line.from}`, qty: line.quantity })),
      ],
      total: stashed.length + sent.reduce((sum, line) => sum + line.quantity, 0),
      warning: [
        blocked.size ? `${plural(blocked.size, 'line')} could not be recorded and are still on the receipt.` : '',
        ...stuck.map(({ device, reason }) => `${machineTitle(device)} is still on the receipt: ${reason}`),
        checklist?.error ? `The return is recorded, but the checklist link could not be made: ${checklist.error}` : '',
      ].filter(Boolean).join(' '),
      next: 'Next return',
      link: { to: '/assets/people', label: 'Open who has what' },
    });
    setBackLines((current) => current.filter((line) => line.handoverId == null || blocked.has(line.handoverId)));
    setMachinesBack(stuck.map(({ device }) => device));
    setMissesBack([]);
  });

  const nextJob = () => {
    if (done?.mode === 'out' && !basket.lines.length && !machines.length) {
      setBasket(newBasket());
      setMissesOut([]);
    }
    setDone(null);
    inputRef.current?.focus();
  };

  // ── Receipt rows, per mode ────────────────────────────────────────────────

  const categories = useMemo(() => categoriesIn(assets), [assets]);

  const draftRows = (theBatch, setTheBatch, theHolds, which) => theBatch.drafts.map((draft) => {
    const unnamed = needsKind(draft);
    const hold = theHolds.get(draft.localId);
    const update = (field) => (event) => setTheBatch((current) => setLineField(current, draft.localId, field, event.target.value));
    const bulk = draft.trackingMode !== TRACKED;
    return {
      id: draft.localId,
      badge: unnamed ? '?' : badgeOf(draft.category),
      name: unnamed ? 'New code'
        : (!String(draft.model ?? '').trim() && !String(draft.manufacturer ?? '').trim() && bulk
          ? `New ${String(draft.category).toLowerCase()} — make and model?`
          : assetTitle(draft)),
      sub: [
        draft.serialNumber ? `S/N ${draft.serialNumber}` : (draft.assetTag ? `Label ${draft.assetTag}` : (draft.partNumber ? `Part ${draft.partNumber}` : '')),
        draft.noSerial && (draft.photoId ? 'No serial · photo taken' : 'No serial · no photo'),
        bulk && serialCount(draft) > 0 && `${serialCount(draft)} serial${serialCount(draft) === 1 ? '' : 's'} recorded`,
        draft.location,
      ].filter(Boolean).join(' · '),
      note: hold,
      tone: hold ? 'ask' : 'ok',
      qty: bulk ? draft.quantity : 1,
      onInc: bulk && !unnamed ? () => setTheBatch((current) => setLineField(current, draft.localId, 'quantity', draft.quantity + 1)) : null,
      onDec: () => setTheBatch((current) => setLineField(current, draft.localId, 'quantity', Math.max(1, draft.quantity - 1))),
      onRemove: () => setTheBatch((current) => removeDraft(current, draft.localId)),
      extra: (unnamed || hold || bulk) && (
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
          {unnamed && (
            <button type="button" className="till-link till-link-small" onClick={() => setSweepFor({ localId: draft.localId, which })}>
              Sweep the box — read what it is off the packaging
            </button>
          )}
          {unnamed && (
            <SameAsSearch
              assets={assets}
              onPick={(asset) => {
                setTheBatch((current) => linkToModel(current, draft.localId, asset));
                say('ok', `Linked to ${asset.title || asset.model} — its box will be recognised from now on`);
              }}
            />
          )}
          {bulk && !unnamed && (
            <button
              type="button"
              className="till-link till-link-small"
              onClick={() => setRunFor({ localId: draft.localId, which, title: `Serials · ${assetTitle(draft)}` })}
            >
              Each has its own serial? Scan them
            </button>
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
  /** What a serial run opens with: the serials and barcodes the line already has. */
  const lineFor = (target) => (target.which === 'count' ? countBatch : batch).drafts.find((draft) => draft.localId === target.localId);
  const serialsOn = (target) => {
    const draft = lineFor(target);
    if (!draft) return [];
    const items = parseUnits(draft.units).map((unit) => unit.serialNumber).filter(Boolean);
    return draft.serialNumber ? [draft.serialNumber, ...items] : items;
  };
  const boxCodesOf = (target) => {
    const draft = lineFor(target);
    return draft ? [draft.partNumber, ...(draft.additionalCodes ?? [])].filter(Boolean) : [];
  };

  const rowsIn = draftRows(batch, setBatch, holds, 'in');
  const rowsCount = draftRows(countBatch, setCountBatch, countHolds, 'count');

  const rowsOut = [
    ...machines.map((device) => ({
      id: `machine:${device.id}`,
      badge: device.deviceType === 'Desktop' ? 'DT' : 'LT',
      name: machineTitle(device),
      sub: machineSub(device),
      tone: 'ok',
      qty: 1,
      onRemove: () => setMachines((list) => list.filter((entry) => entry.id !== device.id)),
    })),
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
    ...machinesBack.map((device) => ({
      id: `machine:${device.id}`,
      badge: '↩',
      name: machineTitle(device),
      sub: [device.serialNumber && `S/N ${device.serialNumber}`, `From ${device.owner || 'unknown owner'} · to the IT Stash`].filter(Boolean).join(' · '),
      tone: 'ok',
      qty: 1,
      onRemove: () => setMachinesBack((list) => list.filter((entry) => entry.id !== device.id)),
    })),
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

  // A checklist is one employee's, so an OUT link is offered only when every
  // returned line came from the same person.
  // A machine knows its owner only by name, so with machines on the receipt
  // the holders are compared by name.
  const oneHolder = (machinesBack.length
    ? new Set([
      ...backLines.filter((line) => line.handoverId != null).map((line) => String(line.from).trim().toLowerCase()),
      ...machinesBack.map((device) => String(device.owner ?? '').trim().toLowerCase()),
    ])
    : new Set(backLines.filter((line) => line.handoverId != null).map((line) => line.fromEmail || line.from))).size === 1;

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
    total = itemsOut(basket, refusals) + machines.length;
    const left = refusals.size + missesOut.length;
    if (left) { note = `${plural(left, 'line')} can’t go out and will be left off`; noteBad = true; }
    cta = !total ? 'Scan what they are getting'
      : !basket.person ? 'Choose who is getting it'
        : `Hand ${total} to ${firstName(basket.person.name)} · sign`;
    ctaOk = total > 0;
    onCheckout = () => setSheet(basket.person ? 'signhow' : 'person');
  } else {
    rows = rowsBack;
    total = itemsBack(backLines) + machinesBack.length;
    const asking = backLines.filter((line) => !line.handoverId).length;
    if (asking) { note = `Say whose ${asking === 1 ? 'it is' : 'they are'} on ${plural(asking, 'line')}`; noteBad = true; }
    cta = total ? `Record return of ${total}` : 'Scan what came back';
    ctaOk = total > 0;
    // A checklist is one employee's: offered only when everything came back from one.
    onCheckout = () => (oneHolder ? setSheet('backhow') : recordReturn(false));
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
    in: itemsIn(batch), out: itemsOut(basket, refusals) + machines.length, back: itemsBack(backLines) + machinesBack.length, count: itemsIn(countBatch),
  };
  const purchase = batch.purchase;
  const setPurchase = (field, value) => setBatch((current) => ({ ...current, purchase: { ...current.purchase, [field]: value } }));

  return (
    <AppShell
      title="Till"
      subtitle="Scan things onto the receipt. One button to finish."
      actions={(
        <Button variant="ghost" icon={ClipboardList} onClick={() => navigate(TILL_CHECKLISTS)}>
          Signed checklists
        </Button>
      )}
    >
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
                  active={!sheet && !runFor && !sweepFor}
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

            {mode === 'out' && stashLeft.length > 0 && (
              <section className="till-keys" aria-label="In the IT Stash">
                <h2 className="till-kicker">In the IT Stash · {stashLeft.length}</h2>
                <ul className="till-outnow">
                  {stashLeft.slice(0, 6).map((device) => (
                    <li key={device.id}>
                      <span>
                        <strong>{machineTitle(device)}</strong>
                        <span>{device.serialNumber ? `S/N ${device.serialNumber}` : (device.deviceType || 'Machine')}</span>
                      </span>
                      <button type="button" className="till-chip" onClick={() => addMachine(device)}>Hand out</button>
                    </li>
                  ))}
                </ul>
                {stashLeft.length > 6 && <p className="till-foot-note">Type a model, name or serial to find the rest.</p>}
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

      {sheet === 'signhow' && basket.person && (
        <ChoiceSheet
          title={`How will ${firstName(basket.person.name)} sign?`}
          lede={machines.length
            ? 'Machines from the IT Stash go to them in the device list. A checklist link is their signed record of the machine.'
            : "A checklist link is the employee's own signed record. Their signature is added to this handover when they sign it."}
          onClose={() => setSheet(null)}
          options={[
            // A phone signature lands on handover rows; a stash machine has
            // none, so a receipt of machines only is signed on a checklist.
            ...(itemsOut(basket, refusals) > 0
              ? [{ key: 'phone', title: 'Sign on this phone', body: machines.length ? 'Hand them the phone now. The machine is not on what they sign — use a checklist link for that.' : 'Hand them the phone now.', onPick: () => setSheet('sign') }]
              : []),
            { key: 'in', title: 'Checklist link · New joiner (IN)', body: 'Items ticked, serials filled in. They sign on the link.', onPick: () => { setSheet(null); handOver(IN); } },
            { key: 'request', title: 'Checklist link · Individual request', body: 'Items listed with quantities, serials filled in. They sign on the link.', onPick: () => { setSheet(null); handOver(INDIVIDUAL); } },
          ]}
        />
      )}

      {sheet === 'backhow' && (
        <ChoiceSheet
          title="Record the return"
          onClose={() => setSheet(null)}
          options={[
            { key: 'plain', title: 'Record the return', body: 'Just put it back in stock.', onPick: () => { setSheet(null); recordReturn(false); } },
            { key: 'out', title: 'Record it and send an OUT checklist', body: 'What came back, ticked and with serials. They sign on the link.', onPick: () => { setSheet(null); recordReturn(true); } },
          ]}
        />
      )}

      {sweepFor && (
        <BoxSweepSheet
          onCancel={() => setSweepFor(null)}
          onUse={(found) => {
            const setTarget = sweepFor.which === 'count' ? setCountBatch : setBatch;
            setTarget((current) => applySweep(current, sweepFor.localId, found));
            setSweepFor(null);
            say('ok', 'Filled in from the box — marked as guessed');
          }}
        />
      )}

      {runFor && (
        <SerialRunSheet
          title={runFor.title}
          assets={assets}
          drafts={(runFor.which === 'count' ? countBatch : batch).drafts}
          existing={serialsOn(runFor)}
          boxCodes={boxCodesOf(runFor)}
          onCancel={() => setRunFor(null)}
          onDone={(run) => {
            const setTarget = runFor.which === 'count' ? setCountBatch : setBatch;
            setTarget((current) => addSerialsTo(current, runFor.localId, run.serials, run.without, run.boxCodes));
            setRunFor(null);
            say('ok', `${run.serials.length + run.without} items recorded`);
          }}
        />
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
          count={itemsOut(basket, refusals)}
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

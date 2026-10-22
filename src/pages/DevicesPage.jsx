import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { useSearchParams } from 'react-router-dom';
import AppShell from '../components/AppShell';
import { gradeCssVars } from '../features/devices/standards/gradeColors';
import { Card, EmptyState, ErrorBanner } from '../components/ui/Surfaces';
import Button from '../components/ui/Button';
import Spinner from '../components/ui/Spinner';
import {
  RefreshCw, ChevronLeft, Check, Settings,
} from '../components/ui/Icons';
import { useSharePointToken } from '../hooks/useRequests';
import DropZone from '../features/devices/ui/DropZone';
import ReviewGrid from '../features/devices/ui/ReviewGrid';
import SaveProgress from '../features/devices/ui/SaveProgress';
import DeviceTable from '../features/devices/ui/DeviceTable';
import DeviceMap from '../features/devices/ui/DeviceMap';
import DeviceFinder from '../features/devices/ui/DeviceFinder';
import ToDoView from '../features/devices/ui/ToDoView';
import StandardsPage from '../features/devices/ui/StandardsPage';
import { importFiles, mergeImports } from '../features/devices/importFiles';
import { issuesFor, sortForReview } from '../features/devices/reviewIssues';
import { useDevices } from '../features/devices/useDevices';
import { useDeviceAdmin } from '../features/devices/useDeviceAdmin';
import { labelOf, IN_FLEET } from '../features/devices/deviceFilters';
import { todoItems } from '../features/devices/stats/todoItems';
import { syncDevices } from '../features/devices/sharepoint/syncDevices';
import { updateDevice, deleteDevice, deleteDevices } from '../features/devices/sharepoint/updateDevice';
import { provisionLists } from '../features/devices/sharepoint/provisionLists';
import { matchIncoming, noticeFor } from '../features/devices/lifecycle/matchIncoming';
import { replacementsFor, unanswered } from '../features/devices/lifecycle/replacements';
import { locationsIn, cleanLocation } from '../features/devices/map/locations';
import { fileNameLocations } from '../features/devices/map/fileNameLocations';
import { inFleet, statusOf, SPARE } from '../features/devices/lifecycle/status';
import { performLifecycle } from '../features/devices/sharepoint/writeLifecycle';
import { ACTIONS } from '../features/devices/lifecycle/planLifecycle';
import { formatMYT } from '../utils/malaysiaTime';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const IDLE_SAVE = {
  phase: 'starting', done: 0, total: 0, results: null, error: null,
  changeCount: 0, unchanged: 0, stintFailures: 0, skipped: [],
};

/** Everything in the query string that is a filter rather than a view switch. */
const FILTER_KEYS = [
  'risk', 'attention', 'type', 'department', 'os', 'av',
  'storage', 'ram', 'cpu', 'windows', 'stale', 'q',
  'part', 'critical', 'persona', 'license', 'server', 'formfit', 'status', 'location', 'work',
];

/** The three views a person switches between. The URL keys are the old tab keys. */
const VIEWS = [
  { key: 'register', label: 'List' },
  { key: 'map', label: 'Places' },
  { key: 'dashboard', label: 'To do' },
];

const IMPORT_STEPS = ['Drop', 'Check', 'Save'];

const plural = (count, word) => `${count} ${word}${count === 1 ? '' : 's'}`;

export default function DevicesPage() {
  const getToken = useSharePointToken();
  const [params, setParams] = useSearchParams();
  const { devices: saved, loading, error, reload, standards } = useDevices();
  // Removing a machine is for admins (owners of the IT helpdesk site) only.
  const isAdmin = useDeviceAdmin();

  const view = params.get('view') ?? 'register';
  const [menuOpen, setMenuOpen] = useState(false);
  const menuRef = useRef(null);
  const menuButtonRef = useRef(null);
  const switchRefs = useRef([]);

  // Close the settings menu on an outside press. Escape is handled on the
  // menu itself, so no page-wide key listener is needed.
  useEffect(() => {
    if (!menuOpen) return undefined;
    const onPress = (event) => {
      if (!menuRef.current?.contains(event.target)) setMenuOpen(false);
    };
    document.addEventListener('mousedown', onPress);
    return () => document.removeEventListener('mousedown', onPress);
  }, [menuOpen]);

  const [stage, setStage] = useState('drop');
  const [parsed, setParsed] = useState([]);
  const [rejected, setRejected] = useState([]);
  const [busy, setBusy] = useState(false);

  // Edits are held apart from the parsed records so that a re-parse or a
  // "start over" discards them cleanly, and the raw record still matches the
  // file it came from.
  const [edits, setEdits] = useState({});
  const [excluded, setExcluded] = useState(new Set());
  const [answers, setAnswers] = useState({});
  const [save, setSave] = useState(IDLE_SAVE);
  const [rowBusy, setRowBusy] = useState(false);
  const [rowError, setRowError] = useState('');
  const columnsChecked = useRef(false);

  const merged = useMemo(
    () => parsed.map((device) => ({ ...device, ...(edits[device.sourceFileName] ?? {}) })),
    [parsed, edits],
  );

  const matches = useMemo(() => matchIncoming(merged, saved), [merged, saved]);
  const prompts = useMemo(() => replacementsFor(matches, saved), [matches, saved]);
  const notices = useMemo(() => new Map(
    matches.map((match) => [match.device.sourceFileName, noticeFor(match)]).filter(([, text]) => text),
  ), [matches]);
  const waiting = unanswered(prompts, answers, excluded);

  const filters = useMemo(
    () => Object.fromEntries(FILTER_KEYS.map((key) => [key, params.get(key) ?? ''])),
    [params],
  );

  // Figures count the machines people are working on. A retired laptop
  // reported as a critical risk would be a figure lying about the fleet.
  const fleet = useMemo(() => saved.filter(inFleet), [saved]);
  const spareCount = useMemo(() => saved.filter((d) => statusOf(d) === SPARE).length, [saved]);
  const newestScan = useMemo(() => {
    let newest = null;
    for (const device of saved) {
      if (typeof device.scannedOn === 'number' && Number.isFinite(device.scannedOn)
        && (newest === null || device.scannedOn > newest)) {
        newest = device.scannedOn;
      }
    }
    return newest;
  }, [saved]);
  const locationOptions = useMemo(() => locationsIn(saved), [saved]);

  // The dashboard reads one department at a time when asked to. It shares the
  // register's `department` key, so a scope chosen here survives the jump into
  // the rows behind any card.
  const department = params.get('department') ?? '';

  const departments = useMemo(() => {
    const names = new Set(fleet.map((device) => labelOf(device.department)));
    return [...names].sort((a, b) => a.localeCompare(b));
  }, [fleet]);

  const scoped = useMemo(
    () => (department
      ? fleet.filter((device) => labelOf(device.department) === department)
      : fleet),
    [fleet, department],
  );

  const todoCount = useMemo(() => todoItems(scoped, Date.now()).length, [scoped]);

  const flagged = merged.filter((device) => issuesFor(device).length > 0).length;
  const included = merged.filter((device) => !excluded.has(device.sourceFileName)).length;

  /**
   * One key, or several at once as an object. Several changes must go out as
   * ONE update: the router's updater reads the address as it was rendered, so
   * a loop of single calls would leave only the last one applied.
   */
  const setParam = useCallback((key, value) => {
    const changes = typeof key === 'object' ? key : { [key]: value };
    setParams((current) => {
      const next = new URLSearchParams(current);
      for (const [name, wanted] of Object.entries(changes)) {
        if (wanted) next.set(name, wanted);
        else next.delete(name);
      }
      return next;
    });
  }, [setParams]);

  /**
   * Switch views. The map's own drill-down keys (location, place) never carry
   * over; `department` is cleared only on the way INTO the map, because it is
   * the map's drill-down there and a plain filter everywhere else.
   */
  const switchView = useCallback((key) => {
    setParams((current) => {
      const next = new URLSearchParams(current);
      next.set('view', key);
      next.delete('location');
      next.delete('place');
      if (key === 'map') next.delete('department');
      return next;
    });
  }, [setParams]);

  const onSwitchKey = (event, position) => {
    if (event.key !== 'ArrowRight' && event.key !== 'ArrowLeft') return;
    event.preventDefault();
    const step = event.key === 'ArrowRight' ? 1 : -1;
    const target = (position + step + VIEWS.length) % VIEWS.length;
    switchRefs.current[target]?.focus();
    switchView(VIEWS[target].key);
  };

  /** Open the register, optionally filtered. `key` null means "no filter". */
  const openRegister = useCallback((key, value, extra) => {
    setParams((current) => {
      const next = new URLSearchParams(current);
      next.set('view', 'register');
      for (const [name, wanted] of Object.entries(extra ?? {})) {
        if (wanted) next.set(name, wanted);
        else next.delete(name);
      }
      // Always default to In use + In repair, but let an explicit status key override it
      if (key !== 'status') next.set('status', IN_FLEET);
      if (!key || key === 'view') return next;
      if (value) {
        next.set(key, value);
      } else {
        next.delete(key);
      }
      return next;
    });
  }, [setParams]);

  const handleFiles = useCallback(async (files) => {
    setBusy(true);
    try {
      const incoming = await importFiles(files, { knownLocations: fileNameLocations(saved) });
      // A second drop adds to the review rather than starting it over, so a
      // batch that arrives in several goes still ends up as one save. Edits
      // already made are keyed by file name and survive untouched.
      const result = parsed.length
        ? mergeImports({ devices: parsed, rejected: [] }, incoming)
        : incoming;
      // Sorted once per drop: the grid must not reorder while somebody edits it.
      const asking = new Set(replacementsFor(matchIncoming(result.devices, saved), saved)
        .map((prompt) => prompt.sourceFileName));
      setParsed(sortForReview(result.devices, asking));
      setRejected(result.rejected);
      if (result.devices.length) setStage('review');
    } finally {
      setBusy(false);
    }
  }, [parsed, saved]);

  const handleChange = (id, key, value) =>
    setEdits((current) => ({ ...current, [id]: { ...(current[id] ?? {}), [key]: value } }));

  const handleToggleRow = (id) =>
    setExcluded((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });

  const resetImport = () => {
    setParsed([]);
    setRejected([]);
    setEdits({});
    setExcluded(new Set());
    setAnswers({});
    setSave(IDLE_SAVE);
    setStage('drop');
  };

  const handleSave = async (onlyNames) => {
    const toSave = merged.filter(
      (device) =>
        !excluded.has(device.sourceFileName)
        && (!onlyNames || onlyNames.includes(device.computerName)),
    );

    setStage('save');
    setSave({ ...IDLE_SAVE, total: toSave.length });

    try {
      const tokenRes = await getToken();
      const outcome = await syncDevices({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        devices: toSave,
        answers,
        changedBy: tokenRes.account?.username ?? '',
        onProgress: ({ phase, done, total }) =>
          setSave((current) => ({ ...current, phase, done, total })),
      });
      setSave((current) => ({
        ...current,
        results: outcome.results,
        changeCount: outcome.changeCount,
        unchanged: outcome.unchanged,
        stintFailures: outcome.stintFailures,
        skipped: outcome.skipped,
      }));
      reload();
    } catch (failure) {
      setSave((current) => ({ ...current, error: failure.message }));
    }
  };

  /** One row edited or removed in the register, rather than a whole import. */
  const runRowAction = async (action) => {
    setRowBusy(true);
    setRowError('');
    try {
      const tokenRes = await getToken();

      // Once per visit, and cheap when there is nothing to do: an edit writes
      // columns a save would have created, so the register must not depend on
      // somebody having run an import first.
      if (!columnsChecked.current) {
        await provisionLists(SHAREPOINT_SITE_URL, tokenRes.accessToken);
        columnsChecked.current = true;
      }

      try {
        await action(tokenRes);
      } finally {
        // The register re-reads whatever the action managed to do, failure or
        // not: removing twelve machines and getting eleven has still changed
        // SharePoint, and a table left showing all twelve would be lying.
        reload();
      }
    } catch (failure) {
      setRowError(failure.message);
    } finally {
      setRowBusy(false);
    }
  };

  const OWNERSHIP = ['owner', 'location', 'department'];
  const trimmed = (value) => String(value ?? '').trim();

  /**
   * An owner, location or department typed into the register is a change of
   * hands, so it goes through the same write as the machine page's Change
   * owner -- the register and the owner history cannot disagree. Clearing the
   * owner still goes the old way: a blank hands the field back to the scan.
   */
  const handleRowSave = (device, edits) => runRowAction(async (tokenRes) => {
    const moved = OWNERSHIP.some((key) => {
      if (!(key in edits)) return false;
      if (key === 'location') return cleanLocation(edits[key]) !== cleanLocation(device[key]);
      // A capitalisation-only edit is not a change of hands; the lifecycle
      // write refuses it, so it goes through the plain edit below.
      return trimmed(edits[key]).toUpperCase() !== trimmed(device[key]).toUpperCase();
    });
    const owner = trimmed('owner' in edits ? edits.owner : device.owner);
    let existing = device;
    let rest = edits;
    // logFailed is deliberately not an error here: the move happened.

    if (moved && owner) {
      // A machine in the stash or the graveyard being given an owner is an
      // assignment: the plain edit would leave a Spare machine with an owner.
      const { plan, pendingStints } = await performLifecycle({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        deviceId: device.id,
        action: inFleet(device) ? ACTIONS.CHANGE_OWNER : ACTIONS.ASSIGN,
        input: {
          owner,
          location: 'location' in edits ? edits.location : device.location,
          department: 'department' in edits ? edits.department : device.department,
        },
        expectedStatus: statusOf(device),
        recordedBy: tokenRes.account?.username ?? '',
      });
      // The remaining edit must start from the manual list the change just
      // wrote, or saving the device type would drop owner from it again.
      existing = { ...device, ...plan.fields };
      rest = Object.fromEntries(Object.entries(edits).filter(([key]) => !OWNERSHIP.includes(key)));
      if (pendingStints.length) {
        throw new Error('The owner changed, but its history could not be written. Open the machine to retry.');
      }
    }

    if (Object.keys(rest).length) {
      await updateDevice({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        existing,
        edits: rest,
        changedBy: tokenRes.account?.username ?? '',
      });
    }
  });

  const handleRowDelete = (device) => runRowAction((tokenRes) => deleteDevice({
    siteUrl: SHAREPOINT_SITE_URL,
    token: tokenRes.accessToken,
    device,
    changedBy: tokenRes.account?.username ?? '',
  }));

  /**
   * A ticked selection removed in one go. `deleteDevices` reports rather than
   * throws, so a machine that would not go has to be named here — the rest of
   * the selection has already gone, and a silent partial removal would leave
   * somebody believing the register is emptier than it is.
   */
  const handleRowDeleteMany = (chosen, onProgress) =>
    runRowAction(async (tokenRes) => {
      const outcome = await deleteDevices({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        devices: chosen,
        changedBy: tokenRes.account?.username ?? '',
        onProgress,
      });

      if (outcome.failures.length) {
        const names = outcome.failures.map((failure) => failure.computerName).join(', ');
        throw new Error(
          `Removed ${outcome.removed.length} of ${chosen.length}. `
          + `Could not remove ${names}: ${outcome.failures[0].error}`,
        );
      }
    });

  const rejectedList = rejected.length > 0 && (
    <ul className="dz-rejected">
      {rejected.map((item) => (
        <li key={item.fileName}>
          <strong>{item.fileName}</strong> — {item.reason}
        </li>
      ))}
    </ul>
  );

  const scopePicker = departments.length > 0 && (
    <label className="dvp-scope">
      <span className="dvp-sr">Department</span>
      <select
        value={department}
        onChange={(event) => setParam('department', event.target.value)}
      >
        <option value="">All departments</option>
        {departments.map((name) => (
          <option key={name} value={name}>{name}</option>
        ))}
      </select>
    </label>
  );

  const stepIndex = { drop: 0, review: 1, save: 2 }[stage] ?? 0;
  const showFinder = view === 'map' || view === 'register' || view === 'dashboard';
  const settingsNote = standards.note || standards.error;
  const factLine = [
    `${fleet.length} in use`,
    `${spareCount} in the IT Stash`,
    newestScan !== null && `newest scan ${formatMYT(newestScan, 'date')}`,
  ].filter(Boolean).join(' · ');

  const actions = (
    <div className="dvp-actions">
      <Button size="md" onClick={() => switchView('import')}>Import scans</Button>
      <div
        className="dvp-menu"
        ref={menuRef}
        onKeyDown={(event) => {
          if (event.key !== 'Escape') return;
          setMenuOpen(false);
          menuButtonRef.current?.focus();
        }}
      >
        <button
          type="button"
          ref={menuButtonRef}
          className="dvp-gear"
          aria-label="Settings"
          aria-haspopup="true"
          aria-expanded={menuOpen}
          onClick={() => setMenuOpen((open) => !open)}
        >
          {loading ? <Spinner size={16} /> : <Settings size={18} />}
          {settingsNote && <span className="dvp-dot" aria-hidden="true" />}
        </button>
        {menuOpen && (
          <div className="dvp-menu-pop">
            {settingsNote && <p className="dvp-menu-note" role="status">{settingsNote}</p>}
            <button
              type="button"
              className="dvp-menu-item"
              onClick={() => { setMenuOpen(false); switchView('standards'); }}
            >
              Grading standard
            </button>
            <button
              type="button"
              className="dvp-menu-item"
              disabled={loading}
              onClick={() => { setMenuOpen(false); reload(); }}
            >
              <RefreshCw size={14} /> Refresh
            </button>
          </div>
        )}
      </div>
    </div>
  );

  const viewSwitch = (
    <div role="group" aria-label="View" className="dvp-switch">
      {VIEWS.map(({ key, label }, position) => {
        const current = view === key;
        return (
          <button
            type="button"
            key={key}
            ref={(node) => { switchRefs.current[position] = node; }}
            aria-current={current ? 'true' : undefined}
            className={`dvp-switch-btn${current ? ' dvp-switch-on' : ''}`}
            onClick={() => switchView(key)}
            onKeyDown={(event) => onSwitchKey(event, position)}
          >
            {label}
            {key === 'dashboard' && <span className="dvp-count">{todoCount}</span>}
          </button>
        );
      })}
    </div>
  );

  return (
    <AppShell title="Devices" subtitle={factLine} actions={actions}>
      <div className="dv-graded dvp-page" style={gradeCssVars(standards.standard.colors)}>
        {showFinder && (
          <div className="dvp-toolbar">
            {/* The List has its own search over the same fields; two boxes would be one too many. */}
            {view !== 'register' && <DeviceFinder devices={saved} />}
            {viewSwitch}
          </div>
        )}

        {error && <ErrorBanner message={error} onRetry={reload} />}

        {view === 'map' && <DeviceMap devices={saved} loading={loading} params={params} standard={standards.standard} />}

        {view === 'dashboard' && (
          <>
            {scopePicker}
            {!loading && scoped.length === 0 ? (
              <Card>
                <EmptyState>
                  {saved.length === 0
                    ? 'Nothing in the register yet. Use Import scans to add your scan reports.'
                    : `No devices are recorded against ${department}.`}
                </EmptyState>
              </Card>
            ) : (
              <ToDoView
                devices={scoped}
                loading={loading}
                spareCount={spareCount}
                onOpen={openRegister}
                onDepartment={(name, part) => openRegister('part', `${part}:Critical`, { department: name })}
              />
            )}
          </>
        )}

        {view === 'register' && (
          <>
            {rowError && <ErrorBanner message={rowError} onRetry={() => setRowError('')} />}
            <DeviceTable
              devices={saved}
              filters={filters}
              onFilterChange={setParam}
              onSave={handleRowSave}
              onDelete={isAdmin ? handleRowDelete : undefined}
              onDeleteMany={isAdmin ? handleRowDeleteMany : undefined}
              busy={rowBusy}
              locationOptions={locationOptions}
            />
          </>
        )}

        {view === 'import' && (
          <section className="dvp-import" aria-labelledby="dvp-import-title">
            <div className="dvp-import-head">
              <button type="button" className="dvp-back" onClick={() => switchView('register')}>
                <ChevronLeft size={14} /> Back to devices
              </button>
              <div className="dvp-import-title-row">
                <h2 id="dvp-import-title" className="dvp-title">Import scans</h2>
                <ol aria-label="Steps" className="dvp-steps">
                  {IMPORT_STEPS.map((label, position) => {
                    const state = position < stepIndex ? 'done' : position === stepIndex ? 'current' : 'future';
                    return (
                      <li
                        key={label}
                        className={`dvp-step dvp-step-${state}`}
                        aria-current={state === 'current' ? 'step' : undefined}
                      >
                        <span className="dvp-step-dot" aria-hidden="true">
                          {state === 'done' ? <Check size={14} /> : position + 1}
                        </span>
                        {label}
                      </li>
                    );
                  })}
                </ol>
              </div>
            </div>

            {stage === 'drop' && (
              <Card className="dz-card">
                <DropZone onFiles={handleFiles} busy={busy} />
                {rejectedList}
              </Card>
            )}

            {stage === 'review' && (
              <Card className="rg-card">
                <div className="dvp-review-bar">
                  <ul className="dvp-figures">
                    <li><strong>{merged.length}</strong> files read</li>
                    <li><strong>{included}</strong> to save</li>
                    <li className={flagged > 0 ? 'dvp-fig-warn' : undefined}>
                      <strong>{flagged}</strong> need attention
                    </li>
                    <li className={waiting > 0 ? 'dvp-fig-amber' : undefined}>
                      <strong>{waiting}</strong> {waiting === 1 ? 'question' : 'questions'}
                    </li>
                  </ul>
                  <div className="dvp-save">
                    {waiting > 0 && (
                      <p className="dvp-save-note" role="status">Answer {plural(waiting, 'question')} first ↓</p>
                    )}
                    <Button variant="secondary" size="sm" onClick={resetImport}>Start over</Button>
                    <Button
                      size="sm"
                      disabled={included === 0 || waiting > 0}
                      onClick={() => handleSave(null)}
                    >
                      Save {plural(included, 'machine')}
                    </Button>
                  </div>
                </div>

                {merged.length === 0 ? (
                  <EmptyState>Nothing to review.</EmptyState>
                ) : (
                  <ReviewGrid
                    devices={merged}
                    excluded={excluded}
                    onChange={handleChange}
                    onToggleRow={handleToggleRow}
                    prompts={prompts}
                    answers={answers}
                    notices={notices}
                    onAnswer={(key, value) => setAnswers((current) => ({ ...current, [key]: value }))}
                  />
                )}

                <DropZone onFiles={handleFiles} busy={busy} compact />

                {rejectedList}
              </Card>
            )}

            {stage === 'save' && (
              <Card>
                <SaveProgress state={save} onRetry={handleSave} onDone={resetImport} />
              </Card>
            )}
          </section>
        )}

        {view === 'standards' && (
          <>
            <button type="button" className="dvp-back" onClick={() => switchView('register')}>
              <ChevronLeft size={14} /> Back to devices
            </button>
            <StandardsPage devices={saved} standards={standards} />
          </>
        )}
      </div>
    </AppShell>
  );
}

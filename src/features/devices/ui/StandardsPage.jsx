import { useMemo, useState } from 'react';
import Button from '../../../components/ui/Button';
import { ErrorBanner } from '../../../components/ui/Surfaces';
import { useConfirm } from '../../../components/ui/useConfirm';
import { useSharePointToken } from '../../../hooks/useRequests';
import { formatMYT } from '../../../utils/malaysiaTime';
import { labelOf, UNASSIGNED } from '../deviceFilters';
import {
  cloneStandard, PROFILE_KEYS, PROFILE_SHORT, GRADES, UNKNOWN, STORAGE_TYPES, DEFAULT_COLORS,
} from '../standards/defaultStandard';
import { validateStandard } from '../standards/validateStandard';
import { diffStandard, summaryOf } from '../standards/diffStandard';
import { previewChanges } from '../standards/previewChanges';
import { inkFor, tooSimilar, GRADE_SLUG } from '../standards/gradeColors';
import { saveStandard } from '../sharepoint/saveStandard';

const SITE = import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const NUMERIC = [
  { part: 'cpu', name: 'CPU generation', unit: 'gen', max: 16, note: 'Intel-equivalent, so AMD lands on the same scale. Pentium, Celeron and AMD before Ryzen are always Critical.' },
  { part: 'ram', name: 'RAM', unit: 'GB', max: 64, note: 'Installed memory (the slots added up), not what Windows reports as usable.' },
  { part: 'storageSize', name: 'Storage size', unit: 'GB', max: 1024, note: 'Total disk, leaving out the scan’s own USB disk. Storage takes the worse of size and type.' },
];
const CUTS = [
  { cut: 'criticalBelow', label: 'Critical below', grade: 'Critical' },
  { cut: 'attentionBelow', label: 'Needs attention below', grade: 'Needs attention' },
  { cut: 'optimalFrom', label: 'Optimal from', grade: 'Optimal' },
];
const CHOICES = [
  { part: 'storageType', name: 'Storage type', note: 'A machine booting from a hard disk waits on it for everything.', rows: STORAGE_TYPES.map((k) => [k, k === 'Mixed' ? 'Mixed (SSD + hard disk)' : k]) },
  { part: 'graphics', name: 'Graphics', note: 'Drawing and rendering need a dedicated card; desk work does not.', rows: [['dedicated', 'Dedicated card'], ['builtIn', 'Built-in graphics']] },
  { part: 'windows', name: 'Windows', note: 'Out of support means no security updates.', rows: [['win11', 'Windows 11'], ['win10Supported', 'Windows 10 (still supported)'], ['outOfSupport', 'Out of support'], ['other', 'Other / older']] },
];

function scaleOf(cut, max) {
  const clamp = (x) => Math.max(0, Math.min(max, Number(x) || 0));
  const points = [0, clamp(cut.criticalBelow), clamp(Math.max(cut.criticalBelow, cut.attentionBelow)),
    clamp(Math.max(cut.attentionBelow, cut.optimalFrom)), max];
  return ['Critical', 'Needs attention', 'Moderate', 'Optimal'].map((grade, i) => ({
    grade, width: `${((points[i + 1] - points[i]) / max) * 100}%`,
  }));
}

/** What counts as Critical … Optimal, per part and profile, and the grade colours. */
export default function StandardsPage({ devices, standards }) {
  const getToken = useSharePointToken();
  const { ask, dialog } = useConfirm();
  const [draft, setDraft] = useState(null);
  const [saving, setSaving] = useState(null);
  const [saveError, setSaveError] = useState('');

  const base = standards.standard;
  const current = draft ?? base;
  const edit = (change) => setDraft((was) => { const next = cloneStandard(was ?? base); change(next); return next; });

  const check = validateStandard(current);
  const errorAt = (path) => check.errors.find((e) => e.path === path)?.message;
  const lines = diffStandard(base, current);
  const preview = useMemo(() => previewChanges(devices, base, current), [devices, base, current]);
  const departments = useMemo(() => [...new Set(devices.map((d) => labelOf(d.department).toUpperCase()))]
    .filter((d) => d !== UNASSIGNED.toUpperCase()).sort(), [devices]);
  const readOnly = !standards.canEdit;

  const save = async (standard, summary, key) => {
    const yes = await ask({
      title: 'Save this standard?',
      body: `${summary}. ${preview.length ? `${preview.reduce((n, p) => n + p.count, 0)} part grades change across the fleet.` : 'No machine changes grade.'}`,
      confirmLabel: 'Save standard',
      cancelLabel: 'Keep editing',
    });
    if (!yes) return;
    setSaving(key);
    setSaveError('');
    try {
      const tokenRes = await getToken();
      await saveStandard({
        siteUrl: SITE, token: tokenRes.accessToken, standard, version: standards.version + 1,
        summary, savedBy: tokenRes.account?.username ?? '',
      });
      setDraft(null);
      standards.reload();
    } catch (failure) {
      setSaveError(failure.message);
    } finally {
      setSaving(null);
    }
  };

  return (
    <section className="sd">
      <header className="sd-head">
        <div>
          <h2 className="dm-title">Device standards</h2>
          <p className="dm-sub">What counts as Critical, Needs attention, Moderate and Optimal, for each part, against the work each profile does. Saving regrades every machine; nothing is re-scanned.</p>
          <p className="dm-sub">
            {standards.version ? `Version ${standards.version} · saved by ${standards.savedBy ?? 'unknown'}${standards.savedOn ? ` on ${formatMYT(standards.savedOn, 'date')}` : ''}` : 'The default standard — nothing saved yet'}
            {readOnly && ' · Only people with edit rights on the IT Device Standards list can change these. Ask IT.'}
          </p>
        </div>
        <div className="sd-actions">
          <Button variant="secondary" disabled={!draft || Boolean(saving)} onClick={() => setDraft(null)}>Discard changes</Button>
          <Button
            loading={saving === 'save'}
            disabled={readOnly || !draft || !check.ok || lines.length === 0}
            onClick={() => save(current, summaryOf(lines), 'save')}
          >
            {check.ok ? `Save standard${lines.length ? ` · ${lines.length} change${lines.length === 1 ? '' : 's'}` : ''}` : 'Fix the order to save'}
          </Button>
        </div>
      </header>

      {saveError && <ErrorBanner message={saveError} />}

      <div className="sd-body">
        <fieldset className="sd-main" disabled={readOnly || Boolean(saving)}>
          <legend className="sr-only">Standard</legend>
          <div className="sd-table">
            <div className="sd-row sd-row-head">
              <span>Part</span>
              {PROFILE_KEYS.map((key) => <span key={key}><strong>{PROFILE_SHORT[key]}</strong><small>{current.profiles[key].label}</small></span>)}
            </div>

            {NUMERIC.map(({ part, name, unit, max, note }) => (
              <div key={part} className="sd-part">
                <div className="sd-part-name"><strong>{name}</strong><small>{note}</small></div>
                {CUTS.map(({ cut, label, grade }) => (
                  <div key={cut} className="sd-row">
                    <span className="sd-row-label"><span className={`sd-dot g-${GRADE_SLUG[grade]}`} />{label}</span>
                    {PROFILE_KEYS.map((key) => {
                      const value = current.profiles[key][part][cut];
                      const changed = value !== base.profiles[key][part][cut];
                      return (
                        <label key={key} className="sd-num">
                          <span className="sr-only">{`${name}, ${PROFILE_SHORT[key]}, ${label}`}</span>
                          <input
                            type="number"
                            min="0"
                            className={changed ? 'sd-changed' : undefined}
                            value={Number.isFinite(value) ? value : ''}
                            onChange={(event) => edit((s) => {
                              s.profiles[key][part][cut] = event.target.value === '' ? null : Number(event.target.value);
                            })}
                          />
                          <span>{unit}</span>
                        </label>
                      );
                    })}
                  </div>
                ))}
                <div className="sd-row">
                  <span className="sd-row-label"><small>How the scale falls</small></span>
                  {PROFILE_KEYS.map((key) => (
                    <div key={key} className="sd-scale">
                      <span className="sd-scale-bar" aria-hidden="true">
                        {scaleOf(current.profiles[key][part], max).map((seg) => (
                          <span key={seg.grade} className={`g-${GRADE_SLUG[seg.grade]}`} style={{ width: seg.width }} />
                        ))}
                      </span>
                      {errorAt(`profiles.${key}.${part}`) && <small role="alert" className="sd-error">{errorAt(`profiles.${key}.${part}`)}</small>}
                    </div>
                  ))}
                </div>
              </div>
            ))}

            {CHOICES.map(({ part, name, note, rows }) => (
              <div key={part} className="sd-part">
                <div className="sd-part-name"><strong>{name}</strong><small>{note}</small></div>
                {rows.map(([kind, label]) => (
                  <div key={kind} className="sd-row">
                    <span className="sd-row-label">{label}</span>
                    {PROFILE_KEYS.map((key) => {
                      const value = current.profiles[key][part][kind];
                      return (
                        <label key={key} className="sd-choice">
                          <span className={`sd-dot g-${GRADE_SLUG[value] ?? 'unknown'}`} aria-hidden="true" />
                          <span className="sr-only">{`${name}, ${label}, ${PROFILE_SHORT[key]}`}</span>
                          <select
                            className={value !== base.profiles[key][part][kind] ? 'sd-changed' : undefined}
                            value={value}
                            onChange={(event) => edit((s) => { s.profiles[key][part][kind] = event.target.value; })}
                          >
                            {GRADES.map((g) => <option key={g} value={g}>{g}</option>)}
                          </select>
                        </label>
                      );
                    })}
                  </div>
                ))}
              </div>
            ))}
          </div>

          <div className="sd-card">
            <h3>Departments</h3>
            <p className="dm-sub">Which profile each department is judged against. Departments not listed, and machines with no department, use Desk.</p>
            <div className="sd-depts">
              {departments.map((department) => (
                <label key={department} className="sd-dept">
                  <span>{department}</span>
                  <select
                    value={current.departments[department] ?? 'desk'}
                    onChange={(event) => edit((s) => { s.departments[department] = event.target.value; })}
                  >
                    {PROFILE_KEYS.map((key) => <option key={key} value={key}>{PROFILE_SHORT[key]}</option>)}
                  </select>
                </label>
              ))}
            </div>
          </div>

          <div className="sd-card">
            <div className="sd-card-head">
              <div>
                <h3>Grade colours</h3>
                <p className="dm-sub">Used for every part chip, bar and badge in the device section. Colour only ever means a grade.</p>
              </div>
              <Button variant="secondary" onClick={() => edit((s) => { s.colors = { ...DEFAULT_COLORS }; })}>Reset to defaults</Button>
            </div>
            <div className="sd-colors">
              {[...GRADES, UNKNOWN].map((grade) => (
                <label key={grade} className="sd-color">
                  <input
                    type="color"
                    value={current.colors[grade]}
                    aria-label={`${grade} colour`}
                    onChange={(event) => edit((s) => { s.colors[grade] = event.target.value; })}
                  />
                  <span><strong>{grade}</strong><small>{current.colors[grade]}</small></span>
                  <span className="pc-chip" style={{ background: current.colors[grade], color: inkFor(current.colors[grade]) }}>RAM · {grade}</span>
                </label>
              ))}
            </div>
            {tooSimilar(current.colors) && <p role="alert" className="sd-warn">Two grade colours are hard to tell apart. You can still save, but people may misread a grade.</p>}
          </div>
        </fieldset>

        <aside className="sd-side">
          <div className="sd-card">
            <h3>If you save now</h3>
            {lines.length === 0 ? (
              <p className="dm-sub">No changes yet. Edit a number or a grade to see which machines would move.</p>
            ) : (
              <>
                <p className="dm-sub">{lines.length} change{lines.length === 1 ? '' : 's'} from {standards.version ? `version ${standards.version}` : 'the default'}.</p>
                <ul className="sd-preview">
                  {preview.map((p) => (
                    <li key={`${p.part}|${p.from}|${p.to}`}>
                      <strong>{p.part.toUpperCase()}</strong> · {p.count} machine{p.count === 1 ? '' : 's'}{' '}
                      <span className={`pc-chip g-${GRADE_SLUG[p.from]}`}>{p.from}</span> → <span className={`pc-chip g-${GRADE_SLUG[p.to]}`}>{p.to}</span>
                    </li>
                  ))}
                  {preview.length === 0 && <li>No machine changes grade.</li>}
                </ul>
              </>
            )}
          </div>

          <div className="sd-card">
            <h3>History</h3>
            {standards.history.length === 0 && <p className="dm-sub">Nothing saved yet — the default standard is in force.</p>}
            <ul className="sd-history">
              {standards.history.map((h) => (
                <li key={h.id}>
                  <span>
                    <strong>v{h.version ?? '?'}</strong> · {h.savedBy ?? 'unknown'}{h.savedOn ? ` · ${formatMYT(h.savedOn, 'date')}` : ''}
                    <small>{h.valid ? h.summary : 'Could not be read'}</small>
                  </span>
                  <Button
                    variant="secondary"
                    size="sm"
                    loading={saving === `restore-${h.id}`}
                    disabled={readOnly || !h.valid || Boolean(saving) || h.version === standards.version}
                    onClick={() => save(h.standard, `Restored version ${h.version}`, `restore-${h.id}`)}
                  >
                    Restore
                  </Button>
                </li>
              ))}
            </ul>
          </div>
        </aside>
      </div>
      {dialog}
    </section>
  );
}

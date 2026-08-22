import { useMemo, useState } from 'react';
import { Card, EmptyState } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import { AlertTriangle, RefreshCw, X } from '../../../components/ui/Icons';
import { useDataStudio } from '../useDataStudio';
import { clusterCategories } from './cleanOps.js';
import { toEpochMs, formatMYT } from '../time/malaysiaTime.js';

// How many dates the ambiguity banner renders both ways. Five is the
// spec's number (§7.5): enough to see the pattern, few enough to read.
const AMBIGUITY_SAMPLE = 5;

// Past this many distinct values a manual merge picker stops being a
// list and starts being a wall. Those columns are free text in practice.
const MANUAL_MERGE_LIMIT = 60;

function confidenceLabel(confidence) {
  if (confidence === 'high') return 'Safe';
  if (confidence === 'medium') return 'Check this';
  return 'Your choice';
}

/**
 * The D/M/Y vs M/D/Y banner (spec §7.5).
 *
 * An ambiguous column defaults to D/M/Y, which is the Malaysian
 * convention and right far more often than not. A conflicting column
 * has no defensible default -- the sheet contains dates that can only
 * be read one way AND dates that can only be read the other -- so it
 * stays unresolved until the user picks.
 *
 * Both readings are rendered side by side rather than described,
 * because "is 05/03 the 5th of March or the 3rd of May" is a question
 * about this data that only this data can answer.
 */
function DateOrderBanner({ column, rawValues }) {
  const { dateOrders, setColumnDateOrder } = useDataStudio();
  const chosen = dateOrders[column.name] ?? (column.dateOrder === 'conflict' ? null : 'dmy');

  const samples = useMemo(() => rawValues
    .filter((v) => typeof v === 'string' && v.trim() !== '')
    .slice(0, AMBIGUITY_SAMPLE)
    .map((raw) => {
      const dmy = toEpochMs(raw, { order: 'dmy', dateOnly: true });
      const mdy = toEpochMs(raw, { order: 'mdy', dateOnly: true });
      return {
        raw,
        dmy: Number.isNaN(dmy) ? '—' : formatMYT(dmy, 'date'),
        mdy: Number.isNaN(mdy) ? '—' : formatMYT(mdy, 'date'),
      };
    }), [rawValues]);

  return (
    <div className="ds-banner" role="region" aria-label={`Date order for ${column.name}`}>
      <div className="ds-banner-head">
        <AlertTriangle size={15} />
        <strong>{column.name}</strong>
        <span>
          {column.dateOrder === 'conflict'
            ? 'has dates written both ways round. Pick which reading is right — nothing is charted until you do.'
            : 'could be read day-first or month-first. Malaysian convention is day-first, which is what is used unless you change it.'}
        </span>
      </div>

      <div className="ds-banner-table-scroll">
        <table className="ds-table ds-banner-table">
          <thead>
            <tr>
              <th>In the sheet</th>
              <th>Day first (D/M/Y)</th>
              <th>Month first (M/D/Y)</th>
            </tr>
          </thead>
          <tbody>
            {samples.map((s, i) => (
              <tr key={`${s.raw}-${i}`}>
                <td className="ds-col-name">{s.raw}</td>
                <td className={chosen === 'dmy' ? 'ds-chosen' : undefined}>{s.dmy}</td>
                <td className={chosen === 'mdy' ? 'ds-chosen' : undefined}>{s.mdy}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="ds-banner-actions">
        <Button
          size="sm"
          variant={chosen === 'dmy' ? 'primary' : 'secondary'}
          onClick={() => setColumnDateOrder(column.name, 'dmy')}
        >
          Read as day first
        </Button>
        <Button
          size="sm"
          variant={chosen === 'mdy' ? 'primary' : 'secondary'}
          onClick={() => setColumnDateOrder(column.name, 'mdy')}
        >
          Read as month first
        </Button>
      </div>
    </div>
  );
}

/**
 * The UTC shift control (spec §9).
 *
 * For a date-ONLY column the control is rendered disabled rather than
 * hidden. Hiding it would leave someone who knows their export is in
 * UTC hunting for a setting that is not there; showing it disabled,
 * with the reason, answers the question they came with.
 */
function ZoneControl({ column }) {
  const { zones, setColumnZone } = useDataStudio();
  const dateOnly = column.type === 'date';
  const value = zones[column.name] ?? 'local';

  return (
    <div className="ds-zone">
      <label className="ds-field">
        <span>{column.name} — stored as</span>
        <select
          className="ds-select"
          value={dateOnly ? 'local' : value}
          disabled={dateOnly}
          onChange={(e) => setColumnZone(column.name, e.target.value)}
        >
          <option value="local">Already Malaysia time</option>
          <option value="utc">UTC — shift by +8 hours</option>
        </select>
      </label>
      {dateOnly && (
        <p className="ds-zone-note">
          This column holds dates with no time of day, so it is never shifted. Adding eight
          hours to a bare date just moves it to the wrong day.
        </p>
      )}
    </div>
  );
}

/** Pick two or more spellings and say which one survives. */
function ManualMerge({ column, rawValues }) {
  const { addManualMerge } = useDataStudio();
  const [picked, setPicked] = useState([]);

  const clusters = useMemo(() => clusterCategories(rawValues), [rawValues]);

  if (clusters.length === 0 || clusters.length > MANUAL_MERGE_LIMIT) return null;

  const toggle = (key) => setPicked((current) => (
    current.includes(key) ? current.filter((k) => k !== key) : [...current, key]));

  const survivor = clusters.find((c) => c.key === picked[0]);

  return (
    <details className="ds-merge">
      <summary>
        Merge categories in {column.name} by hand ({clusters.length} distinct)
      </summary>

      <p className="ds-merge-hint">
        Tick two or more. The first one you tick is the spelling that survives.
      </p>

      <ul className="ds-merge-list">
        {clusters.map((cluster) => (
          <li key={cluster.key}>
            <label>
              <input
                type="checkbox"
                checked={picked.includes(cluster.key)}
                onChange={() => toggle(cluster.key)}
              />
              <span className="ds-merge-label">{cluster.canonical}</span>
              <span className="ds-merge-count">{cluster.count}</span>
            </label>
          </li>
        ))}
      </ul>

      <div className="ds-merge-actions">
        <Button
          size="sm"
          disabled={picked.length < 2 || !survivor}
          onClick={() => {
            addManualMerge(column.name, picked, survivor.canonical);
            setPicked([]);
          }}
        >
          {survivor && picked.length >= 2
            ? `Merge ${picked.length} into "${survivor.canonical}"`
            : 'Merge selected'}
        </Button>
        {picked.length > 0 && (
          <Button size="sm" variant="secondary" onClick={() => setPicked([])}>
            Clear
          </Button>
        )}
      </div>
    </details>
  );
}

function StepRow({ step }) {
  const { setStepEnabled, removeStep } = useDataStudio();

  return (
    <li className={`ds-step ds-step-${step.confidence}`}>
      <label className="ds-step-main">
        <input
          type="checkbox"
          checked={step.enabled}
          onChange={(e) => setStepEnabled(step.id, e.target.checked)}
        />
        <span className="ds-step-text">{step.preview}</span>
      </label>
      <span className="ds-step-count">{step.affectedCount.toLocaleString()}</span>
      <span className={`ds-badge ds-badge-${step.confidence}`}>
        {confidenceLabel(step.confidence)}
      </span>
      {step.manual && (
        <button
          type="button"
          className="ds-step-remove"
          aria-label={`Remove the manual merge on ${step.column}`}
          onClick={() => removeStep(step.id)}
        >
          <X size={13} />
        </button>
      )}
    </li>
  );
}

export default function CleanReview() {
  const {
    profile, grid, plan, dataset, cleaning, commitClean, setStage, textColumns, startAnalysis,
  } = useDataStudio();

  const rawByColumn = useMemo(() => {
    if (!profile || !grid) return new Map();
    return new Map(profile.columns.map((c) => [
      c.name, grid.rows.map((row) => row?.[c.index]),
    ]));
  }, [profile, grid]);

  // Grouped by column so the checklist reads as "here is what happens to
  // Department", rather than as a flat list the user has to re-sort in
  // their head. Whole-grid steps get their own group at the end.
  const groups = useMemo(() => {
    const byColumn = new Map();
    const wholeGrid = [];
    for (const step of plan) {
      if (step.column === null) {
        wholeGrid.push(step);
        continue;
      }
      if (!byColumn.has(step.column)) byColumn.set(step.column, []);
      byColumn.get(step.column).push(step);
    }
    return { byColumn, wholeGrid };
  }, [plan]);

  const needsDateChoice = useMemo(
    () => (profile?.columns ?? []).filter(
      (c) => c.role === 'temporal'
        && (c.dateOrder === 'ambiguous' || c.dateOrder === 'conflict'),
    ),
    [profile],
  );

  const temporal = useMemo(
    () => (profile?.columns ?? []).filter((c) => c.role === 'temporal'),
    [profile],
  );

  const dimensions = useMemo(
    () => (profile?.columns ?? []).filter((c) => c.role === 'dimension'),
    [profile],
  );

  if (!profile || !grid) return <EmptyState>Nothing imported yet.</EmptyState>;

  const enabledCount = plan.filter((s) => s.enabled).length;

  return (
    <>
      <div className="ds-toolbar">
        <span className="ds-summary">
          {plan.length === 0
            ? 'Nothing needs cleaning — this sheet came in tidy.'
            : `${enabledCount} of ${plan.length} steps will run`}
        </span>
        <span className="ds-toolbar-spacer" />
        <span className="ds-summary">
          {dataset
            ? `${dataset.rowCount.toLocaleString()} rows · ${dataset.columns.length} columns`
            : 'Working…'}
        </span>
        <Button variant="secondary" size="sm" icon={RefreshCw} onClick={() => setStage('profiled')}>
          Back to columns
        </Button>
        {/* Only when the sheet holds written answers worth reading. */}
        {textColumns.length > 0 && (
          <Button
            variant="secondary"
            size="sm"
            onClick={() => startAnalysis(textColumns[0].name)}
          >
            Text analysis
          </Button>
        )}
        <Button size="sm" disabled={!dataset || cleaning} onClick={commitClean}>
          Build the dashboard
        </Button>
      </div>

      {needsDateChoice.map((column) => (
        <DateOrderBanner
          key={column.name}
          column={column}
          rawValues={rawByColumn.get(column.name) ?? []}
        />
      ))}

      {plan.length === 0 ? (
        <Card>
          <EmptyState>
            No padding, no placeholder text, no duplicate rows, nothing to merge. Go straight
            to the dashboard.
          </EmptyState>
        </Card>
      ) : (
        <Card className="ds-plan-card">
          {[...groups.byColumn.entries()].map(([columnName, steps]) => (
            <section key={columnName} className="ds-plan-group">
              <h3 className="ds-plan-heading">{columnName}</h3>
              <ul className="ds-step-list">
                {steps.map((step) => <StepRow key={step.id} step={step} />)}
              </ul>
            </section>
          ))}

          {groups.wholeGrid.length > 0 && (
            <section className="ds-plan-group">
              <h3 className="ds-plan-heading">The whole sheet</h3>
              <ul className="ds-step-list">
                {groups.wholeGrid.map((step) => <StepRow key={step.id} step={step} />)}
              </ul>
            </section>
          )}
        </Card>
      )}

      {temporal.length > 0 && (
        <Card className="ds-plan-card">
          <h3 className="ds-plan-heading">Time zones</h3>
          {temporal.map((column) => <ZoneControl key={column.name} column={column} />)}
        </Card>
      )}

      {dimensions.length > 0 && (
        <Card className="ds-plan-card">
          <h3 className="ds-plan-heading">Merge by hand</h3>
          {dimensions.map((column) => (
            <ManualMerge
              key={column.name}
              column={column}
              rawValues={rawByColumn.get(column.name) ?? []}
            />
          ))}
        </Card>
      )}
    </>
  );
}

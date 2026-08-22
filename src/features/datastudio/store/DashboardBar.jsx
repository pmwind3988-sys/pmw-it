import { useRef, useState } from 'react';
import Button from '../../../components/ui/Button';
import { ErrorBanner } from '../../../components/ui/Surfaces';
import { Trash2, Download, Plus } from '../../../components/ui/Icons';
import { useDataStudio } from '../useDataStudio';
import { useDashboards } from './useDashboards.js';
import {
  exportDatasetCsv, exportDashboardJson, importDashboardJson,
} from './exporters.js';

/**
 * Save, load, export and import dashboards for the current dataset.
 *
 * Saving a dataset is a separate action from saving a dashboard on
 * purpose: the dataset is the expensive thing and the dashboard is a
 * view of it, so someone can keep several dashboards over one import
 * without three copies of the data.
 */
export default function DashboardBar() {
  const {
    datasetId, dataset, tiles, globalFilters, fileName,
    saveCurrentDataset, applyDashboard,
  } = useDataStudio();

  const { dashboards, error, save, remove } = useDashboards(datasetId);
  const [name, setName] = useState('');
  const [notice, setNotice] = useState('');
  const [busy, setBusy] = useState(false);
  const fileRef = useRef(null);

  const handleSaveDashboard = async () => {
    setBusy(true);
    // A dashboard has to belong to a saved dataset, or reloading the
    // page would leave it pointing at nothing. Saving the dataset first
    // is invisible and always what the user meant.
    const id = datasetId ?? await saveCurrentDataset(fileName);
    if (id) await save(name.trim() || 'Dashboard', tiles, globalFilters);
    setName('');
    setBusy(false);
  };

  const handleImport = async (file) => {
    const result = await importDashboardJson(file, dataset);
    if (!result.ok) {
      setNotice(result.reason);
      return;
    }
    applyDashboard(result.dashboard);
    setNotice(
      result.missingColumns.length > 0
        ? `Loaded, but this data has no ${result.missingColumns.join(', ')} — those tiles will say so.`
        : `Loaded "${result.dashboard.name}".`,
    );
  };

  return (
    <div className="ds-dashboardbar">
      {error && <ErrorBanner message={error} />}
      {notice && <p className="ds-summary">{notice}</p>}

      <div className="ds-dashboardbar-row">
        <label className="ds-field">
          <span>Save this dashboard as</span>
          <input
            className="ds-select"
            value={name}
            placeholder="Dashboard name"
            onChange={(e) => setName(e.target.value)}
          />
        </label>
        <Button size="sm" disabled={busy || tiles.length === 0} onClick={handleSaveDashboard}>
          Save
        </Button>

        {dashboards.length > 0 && (
          <label className="ds-field">
            <span>Open a saved dashboard</span>
            <select
              className="ds-select"
              value=""
              onChange={(e) => {
                const found = dashboards.find((d) => d.id === e.target.value);
                if (found) applyDashboard(found);
              }}
            >
              <option value="">Pick one…</option>
              {dashboards.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}
            </select>
          </label>
        )}

        <span className="ds-toolbar-spacer" />

        <Button
          size="sm"
          variant="secondary"
          icon={Download}
          disabled={!dataset}
          onClick={() => exportDatasetCsv(dataset, fileName || 'data')}
        >
          CSV
        </Button>
        <Button
          size="sm"
          variant="secondary"
          icon={Download}
          disabled={tiles.length === 0}
          onClick={() => exportDashboardJson(
            { name: name.trim() || 'Dashboard', tiles, globalFilters, datasetName: fileName },
            dataset,
            name.trim() || 'dashboard',
          )}
        >
          Dashboard file
        </Button>
        <Button size="sm" variant="secondary" icon={Plus} onClick={() => fileRef.current?.click()}>
          Load a dashboard file
        </Button>
        <input
          ref={fileRef}
          type="file"
          accept=".json,application/json"
          className="ds-drop-input"
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) handleImport(file);
            e.target.value = '';
          }}
        />
      </div>

      {dashboards.length > 0 && (
        <ul className="ds-dashboard-list">
          {dashboards.map((d) => (
            <li key={d.id}>
              <span>{d.name}</span>
              <button
                type="button"
                className="ds-step-remove"
                aria-label={`Delete the saved dashboard ${d.name}`}
                onClick={() => remove(d.id)}
              >
                <Trash2 size={13} />
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

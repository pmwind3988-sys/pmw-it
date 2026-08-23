import { useMemo, useState } from 'react';
import { Link, useNavigate, useSearchParams } from 'react-router-dom';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import StatCard from '../components/ui/StatCard';
import { Card, ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import {
  ScanLine, Plus, Package, Boxes, Tag, AlertTriangle, Truck, RefreshCw,
} from '../components/ui/Icons';
import { useAssets } from '../features/assets/useAssets';
import { useBatches } from '../features/assets/useBatches';
import { assetStats } from '../features/assets/stats/assetStats';
import {
  filterAssets, sortAssets, optionsFor,
} from '../features/assets/assetFilters';
import { CATEGORIES, CONDITIONS, STATUSES } from '../features/assets/assetKinds';
import { batchTitle } from '../features/assets/draft/batch';
import { newBatch } from '../features/assets/draft/batch';
import { saveBatch } from '../features/assets/store/assetDb';
import AssetTable from '../features/assets/ui/AssetTable';

/**
 * The register: what IT owns, and the two ways to add to it.
 *
 * Filters live in the query string, the same as `/requests` and `/devices`, so
 * a narrowed view is a link somebody can send.
 */
export default function AssetsPage() {
  const navigate = useNavigate();
  const { assets, loading, error, reload } = useAssets();
  const { batches, pendingItems, discard } = useBatches();
  const [params, setParams] = useSearchParams();
  const [query, setQuery] = useState(params.get('q') ?? '');

  const filters = {
    query,
    category: params.get('category') ?? '',
    status: params.get('status') ?? '',
    condition: params.get('condition') ?? '',
    location: params.get('location') ?? '',
    unlabelled: params.get('unlabelled') === '1',
  };

  const setFilter = (key, value) => {
    const next = new URLSearchParams(params);
    if (value) next.set(key, value);
    else next.delete(key);
    setParams(next, { replace: true });
  };

  const stats = useMemo(() => assetStats(assets), [assets]);
  const shown = useMemo(
    () => sortAssets(filterAssets(assets, filters)),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [assets, query, params],
  );

  /** Adding by hand is the same review grid, just with nothing scanned into it. */
  const addByHand = async () => {
    const batch = newBatch();
    await saveBatch(batch);
    navigate(`/assets/batch/${batch.id}`);
  };

  return (
    <AppShell
      title="Asset inventory"
      subtitle="Everything IT owns — scanned in, labelled, and counted"
      search={{ value: query, onChange: setQuery, placeholder: 'Serial, model, label…' }}
      actions={(
        <>
          <Button variant="secondary" icon={Plus} onClick={addByHand}>Add by hand</Button>
          <Button icon={ScanLine} onClick={() => navigate('/assets/scan')}>Scan a delivery</Button>
        </>
      )}
    >
      {/* Not dismissible on purpose: an unsaved batch is invisible to everybody
          but this device, and the only thing standing between a scanned
          delivery and being forgotten is this line (§4.1). */}
      {batches.length > 0 && (
        <Card className="as-notice as-notice-warn">
          <Truck size={16} />
          <span>
            {batches.length} deliver{batches.length === 1 ? 'y is' : 'ies are'} still on this
            device, holding {pendingItems} item{pendingItems === 1 ? '' : 's'}. Nobody else can
            see them until they are saved.
          </span>
          <span className="as-notice-links">
            {batches.map((batch) => (
              <Link key={batch.id} to={`/assets/batch/${batch.id}`} className="as-link">
                {batchTitle(batch)} ({batch.drafts?.length ?? 0})
              </Link>
            ))}
          </span>
        </Card>
      )}

      {error && <ErrorBanner message={error} onRetry={reload} />}

      <div className="stat-grid">
        <StatCard icon={Package} label="Items owned" value={stats.units} loading={loading} />
        <StatCard
          icon={Boxes}
          label="Tracked units"
          value={stats.trackedUnits}
          loading={loading}
          onClick={() => setFilter('trackingMode', 'Tracked')}
        />
        <StatCard
          icon={Tag}
          label="Waiting for a label"
          value={stats.unlabelled}
          loading={loading}
          onClick={() => setFilter('unlabelled', filters.unlabelled ? '' : '1')}
        />
        <StatCard
          icon={AlertTriangle}
          label="Faulty"
          value={stats.faulty}
          loading={loading}
          onClick={() => setFilter('condition', 'Faulty')}
        />
      </div>

      <Card className="as-filters">
        <select
          value={filters.category}
          onChange={(e) => setFilter('category', e.target.value)}
          aria-label="Category"
        >
          <option value="">Every category</option>
          {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>

        <select
          value={filters.status}
          onChange={(e) => setFilter('status', e.target.value)}
          aria-label="Status"
        >
          <option value="">Any status</option>
          {STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
        </select>

        <select
          value={filters.condition}
          onChange={(e) => setFilter('condition', e.target.value)}
          aria-label="Condition"
        >
          <option value="">Any condition</option>
          {CONDITIONS.map((c) => <option key={c} value={c}>{c}</option>)}
        </select>

        <select
          value={filters.location}
          onChange={(e) => setFilter('location', e.target.value)}
          aria-label="Location"
        >
          <option value="">Anywhere</option>
          {optionsFor(assets, 'location').map((l) => <option key={l} value={l}>{l}</option>)}
        </select>

        <Button variant="ghost" size="sm" icon={RefreshCw} onClick={reload}>Refresh</Button>
      </Card>

      {loading && <div className="spinner" />}

      {!loading && shown.length === 0 && (
        <EmptyState>
          {assets.length === 0
            ? 'Nothing in the register yet. Scan a delivery, or add an item by hand.'
            : 'Nothing matches those filters.'}
        </EmptyState>
      )}

      {shown.length > 0 && (
        <>
          <p className="as-hint">
            {shown.length} of {assets.length} row{assets.length === 1 ? '' : 's'}
          </p>
          <AssetTable assets={shown} />
        </>
      )}

      {batches.length > 0 && (
        <details className="as-batchlist">
          <summary>Deliveries waiting on this device</summary>
          <ul>
            {batches.map((batch) => (
              <li key={batch.id}>
                <Link to={`/assets/batch/${batch.id}`} className="as-link">
                  {batchTitle(batch)}
                </Link>
                <span className="as-sub">{batch.drafts?.length ?? 0} items</span>
                <Button variant="ghost" size="sm" onClick={() => discard(batch.id)}>Discard</Button>
              </li>
            ))}
          </ul>
        </details>
      )}
    </AppShell>
  );
}

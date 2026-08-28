import { useMemo, useState } from 'react';
import { Link, useNavigate, useSearchParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import StatCard from '../components/ui/StatCard';
import { Card, ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import {
  ScanLine, Plus, Package, Tag, AlertTriangle, Truck, RefreshCw, Users, Clock,
  ClipboardList, Boxes, X,
} from '../components/ui/Icons';
import { useAssets, SHAREPOINT_SITE_URL } from '../features/assets/useAssets';
import { useSharePointToken } from '../hooks/useRequests';
import { blockersFor, planCombine } from '../features/assets/combine';
import { combineAssets } from '../features/assets/sharepoint/combineAssets';
import { useBatches } from '../features/assets/useBatches';
import { useHandovers } from '../features/assets/useHandovers';
import { isOverdue } from '../features/assets/handover/availability';
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
  const { handovers } = useHandovers();
  const [params, setParams] = useSearchParams();
  const [query, setQuery] = useState(params.get('q') ?? '');

  const { instance } = useMsal();
  const getToken = useSharePointToken();
  // Rows that are really one thing bought ten times, being put back together.
  // Off unless asked for: a register you cannot read without ticking things by
  // accident is worse than one that takes an extra tap to tidy.
  const [picking, setPicking] = useState(false);
  const [picked, setPicked] = useState([]);
  const [combining, setCombining] = useState(false);
  const [combined, setCombined] = useState('');
  const [combineError, setCombineError] = useState('');

  const filters = {
    query,
    category: params.get('category') ?? '',
    status: params.get('status') ?? '',
    condition: params.get('condition') ?? '',
    location: params.get('location') ?? '',
    unlabelled: params.get('unlabelled') === '1',
    pending: params.get('pending') === '1',
  };

  const setFilter = (key, value) => {
    const next = new URLSearchParams(params);
    if (value) next.set(key, value);
    else next.delete(key);
    setParams(next, { replace: true });
  };

  const stats = useMemo(() => assetStats(assets), [assets]);
  const overdue = useMemo(
    () => handovers.filter((row) => isOverdue(row)).length,
    [handovers],
  );
  const shown = useMemo(
    () => sortAssets(filterAssets(assets, filters)),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [assets, query, params],
  );

  const chosen = useMemo(
    () => assets.filter((row) => picked.includes(row.id)),
    [assets, picked],
  );
  const blockers = picked.length ? blockersFor(chosen, handovers) : [];
  const plan = picked.length >= 2 && !blockers.length ? planCombine(chosen) : null;

  const togglePick = (id) => setPicked(
    (current) => (current.includes(id)
      ? current.filter((other) => other !== id)
      : [...current, id]),
  );

  const stopPicking = () => {
    setPicking(false);
    setPicked([]);
    setCombineError('');
  };

  /**
   * Ten rows become one line of ten. It asks first: the rows themselves go,
   * and while everything typed on them is carried onto the items of the
   * surviving line, undoing it means retyping ten rows.
   */
  const combine = async () => {
    const summary = `Combine ${chosen.length} rows into one line of `
      + `${plan.edits.quantity}? Everything on them is kept as ${plan.edits.quantity} `
      + `items on "${plan.keep.title || plan.keep.model || 'the oldest row'}", `
      + `and the other ${plan.remove.length} rows are removed.`;
    if (!window.confirm(summary)) return;

    setCombining(true);
    setCombineError('');
    try {
      const tokenRes = await getToken();
      const account = instance.getActiveAccount();
      const result = await combineAssets({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        rows: chosen,
        changedBy: account?.username ?? account?.name ?? '',
      });

      setCombined(
        `Combined into one line of ${result.quantity}.`
        + (result.failures.length
          ? ` ${result.failures.length} of the old rows could not be removed — `
            + 'everything is safe, try combining what is left again.'
          : ''),
      );
      stopPicking();
      reload();
    } catch (failure) {
      setCombineError(failure.message || 'Those rows could not be combined');
    } finally {
      setCombining(false);
    }
  };

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
          <Button
            variant="ghost"
            icon={picking ? X : Boxes}
            onClick={() => (picking ? stopPicking() : setPicking(true))}
          >
            {picking ? 'Stop combining' : 'Combine rows'}
          </Button>
          <Button variant="ghost" icon={Plus} onClick={addByHand}>Add by hand</Button>
          <Button variant="secondary" icon={Users} onClick={() => navigate('/assets/handover')}>
            Hand over
          </Button>
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
          icon={Users}
          label="Out with people"
          value={stats.out}
          loading={loading}
          onClick={() => navigate('/assets/people')}
        />
        <StatCard
          icon={Clock}
          label="Overdue"
          value={overdue}
          loading={loading}
          onClick={() => navigate('/assets/people')}
        />
        <StatCard
          icon={Tag}
          label="Waiting for a label"
          value={stats.unlabelled}
          loading={loading}
          onClick={() => setFilter('unlabelled', filters.unlabelled ? '' : '1')}
        />
        <StatCard
          icon={ClipboardList}
          label="Needs details"
          value={stats.pending}
          loading={loading}
          onClick={() => setFilter('pending', filters.pending ? '' : '1')}
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

      {/* Ten rows of one monitor that were always one line of ten. What it is
          about to do is spelled out before it does it, because the rows
          themselves go and only the items on the survivor remain. */}
      {picking && (
        <Card className="as-notice as-combine">
          <Boxes size={16} />
          <span>
            {picked.length < 2
              ? 'Tick the rows that are the same thing, and they become one line '
                + 'with each of them kept as an item on it.'
              : `${picked.length} rows → one line of ${plan?.edits.quantity ?? '—'}.`}
            {plan?.warnings.length > 0 && (
              <strong>
                {' '}They do not agree on {plan.warnings.join(' or ')} — check that
                these really are the same thing.
              </strong>
            )}
            {blockers.map((blocker) => <strong key={blocker}> {blocker}</strong>)}
          </span>
          <Button
            size="sm"
            disabled={!plan || combining}
            onClick={combine}
          >
            {combining ? 'Combining…' : 'Combine into one line'}
          </Button>
        </Card>
      )}

      {combined && !picking && (
        <Card className="as-notice as-notice-ok">
          <Boxes size={16} />
          <span>{combined}</span>
        </Card>
      )}

      {combineError && <ErrorBanner message={combineError} onRetry={() => setCombineError('')} />}

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
          <AssetTable
            assets={shown}
            picking={picking}
            picked={picked}
            onPick={togglePick}
          />
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

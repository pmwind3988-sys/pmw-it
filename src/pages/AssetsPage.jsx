import { useMemo, useState } from 'react';
import { Link, useNavigate, useSearchParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import StatCard from '../components/ui/StatCard';
import { Card, ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import Collapsible from '../components/ui/Collapsible';
import Pager from '../components/ui/Pager';
import { paginate } from '../components/ui/paginate';
import { useConfirm } from '../components/ui/useConfirm';
import {
  ScanLine, Plus, Package, Tag, AlertTriangle, Truck, RefreshCw, Users, Clock,
  ClipboardList, Boxes, X,
} from '../components/ui/Icons';
import { useAssets, SHAREPOINT_SITE_URL } from '../features/assets/useAssets';
import { useSharePointToken } from '../hooks/useRequests';
import { blockersFor, planCombine, stillOut } from '../features/assets/combine';
import { combineAssets } from '../features/assets/sharepoint/combineAssets';
import { useBatches } from '../features/assets/useBatches';
import { useHandovers } from '../features/assets/useHandovers';
import { isOverdue } from '../features/assets/handover/availability';
import { assetStats } from '../features/assets/stats/assetStats';
import {
  filterAssets, sortAssets, optionsFor,
} from '../features/assets/assetFilters';
import { CONDITIONS, STATUSES } from '../features/assets/assetKinds';
import { categoriesIn } from '../features/assets/categories';
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
  const { ask, dialog } = useConfirm();

  // Only the page being looked at is built. Two thousand rows laid out at once
  // is a phone that appears to have hung.
  const [pageSize, setPageSize] = useState(25);
  const [at, setAt] = useState({ of: '', page: 1 });

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

  /**
   * Back to the first page whenever the list underneath changes. Staying on
   * page 7 of a search with three results shows an empty table, which reads as
   * "nothing found".
   *
   * The page number is held together with the list it counts rather than being
   * reset by an effect: an effect would render the wrong page first and correct
   * it afterwards, which is a visible flicker of somebody else's rows.
   */
  const of = `${query}|${params.toString()}|${pageSize}`;
  const page = at.of === of ? at.page : 1;
  const setPage = (next) => setAt({ of, page: next });

  const paged = useMemo(() => paginate(shown, page, pageSize), [shown, page, pageSize]);

  // The built-in kinds, plus anything the register is actually using — which
  // is how a category somebody added shows up here without a second list.
  const categories = useMemo(() => categoriesIn(assets), [assets]);

  /** "category, make or model" rather than "category or make or model". */
  const listOf = (parts) => (parts.length < 2
    ? parts.join('')
    : `${parts.slice(0, -1).join(', ')} or ${parts[parts.length - 1]}`);

  /** What the filter panel would say, for the line it shows while folded. */
  const activeFilters = [
    filters.category,
    filters.status,
    filters.condition,
    filters.location,
    filters.unlabelled && 'waiting for a label',
    filters.pending && 'needs details',
  ].filter(Boolean);

  const chosen = useMemo(
    () => assets.filter((row) => picked.includes(row.id)),
    [assets, picked],
  );
  const blockers = picked.length ? blockersFor(chosen) : [];
  const plan = picked.length >= 2 && !blockers.length ? planCombine(chosen, handovers) : null;
  // What is on somebody's desk right now. Not a refusal any more: it stays
  // out, and its handover follows it onto the combined line.
  const outNow = picked.length ? stillOut(chosen, handovers) : [];
  const holders = new Set(outNow.map((row) => row.personEmail || row.personName)).size;

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
    const summary = `Everything on them is kept as ${plan.edits.quantity} `
      + `items on "${plan.keep.title || plan.keep.model || 'the oldest row'}", `
      + `and the other ${plan.remove.length} rows are removed. This cannot be undone.`
      + (outNow.length
        ? ` The ${outNow.length} still out stay out — the same people keep the same `
          + 'things, as items on the new line.'
        : '');
    if (!await ask({
      title: `Combine ${chosen.length} rows?`,
      body: summary,
      confirmLabel: 'Combine into one line',
      cancelLabel: 'Leave them as they are',
    })) return;

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
        + (result.moved
          ? ` ${result.moved} handover${result.moved === 1 ? '' : 's'} now point`
            + `${result.moved === 1 ? 's' : ''} at it, so nobody's item changed hands.`
          : '')
        + (result.failures.length
          ? ` ${result.failures.length} of the old rows could not be finished — `
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

  /**
   * A delivery still sitting on this phone. Nobody else can see it, so
   * discarding it is the one deletion here with no copy anywhere — it asks.
   */
  const discardBatch = async (batch) => {
    const held = batch.drafts?.length ?? 0;
    if (!await ask({
      title: 'Discard this delivery?',
      body: (held
        ? `"${batchTitle(batch)}" holds ${held} item${held === 1 ? '' : 's'} that `
          + `${held === 1 ? 'has' : 'have'} never been saved. Nobody else can see `
          + `${held === 1 ? 'it' : 'them'}, so nothing is left behind once this goes.`
        : `"${batchTitle(batch)}" holds no items yet, so nothing is lost — the `
          + 'delivery details themselves go.'),
      confirmLabel: 'Discard it',
      cancelLabel: 'Keep it',
    })) return;
    discard(batch.id);
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

      {/* Five dropdowns fill a phone screen before a single row of the
          register is visible. They fold away, and the header says what is on
          while they are folded — a hidden filter is a page lying about what it
          is showing. */}
      <Collapsible
        id="assets-filters"
        title="Filters"
        defaultOpen={false}
        summary={activeFilters.length
          ? activeFilters.join(' · ')
          : 'Everything, anywhere, any condition'}
        actions={(
          <Button variant="ghost" size="sm" icon={RefreshCw} onClick={reload}>Refresh</Button>
        )}
      >
        <div className="as-filters">
          <select
            value={filters.category}
            onChange={(e) => setFilter('category', e.target.value)}
            aria-label="Category"
          >
            <option value="">Every category</option>
            {categories.map((c) => <option key={c} value={c}>{c}</option>)}
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
            aria-label="Anywhere"
          >
            <option value="">Anywhere</option>
            {optionsFor(assets, 'location').map((l) => <option key={l} value={l}>{l}</option>)}
          </select>
        </div>
      </Collapsible>

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
                {' '}They do not agree on {listOf(plan.warnings)} — check that
                these really are the same thing.
              </strong>
            )}
            {outNow.length > 0 && (
              <>
                {' '}{outNow.length} item{outNow.length === 1 ? ' is' : 's are'} out with{' '}
                {holders} {holders === 1 ? 'person' : 'people'}, and stay
                {outNow.length === 1 ? 's' : ''} out — the same{' '}
                {holders === 1 ? 'person keeps' : 'people keep'} the same{' '}
                {outNow.length === 1 ? 'thing' : 'things'}, as items on the one line.
              </>
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
            assets={paged.rows}
            picking={picking}
            picked={picked}
            onPick={togglePick}
          />
          <Pager
            page={paged}
            onPage={setPage}
            size={pageSize}
            onSize={setPageSize}
            label="rows"
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
                <Button variant="ghost" size="sm" onClick={() => discardBatch(batch)}>Discard</Button>
              </li>
            ))}
          </ul>
        </details>
      )}

      {dialog}
    </AppShell>
  );
}

import { useEffect, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import AppShell from '../components/AppShell';
import StatCard from '../components/ui/StatCard';
import Button from '../components/ui/Button';
import { Card, EmptyState, ErrorBanner } from '../components/ui/Surfaces';
import { RequestTypeBadge } from '../components/ui/Badges';
import { BarChart, ColumnChart } from '../components/ui/Charts';
import {
  ClipboardList,
  UserPlus,
  UserMinus,
  Calendar,
  Clock,
  Laptop,
  RefreshCw,
  Plus,
  ChevronRight,
} from '../components/ui/Icons';
import { useRequests, requestDate, formatDate, toChoiceArray } from '../hooks/useRequests';
import { useSession } from '../hooks/useSession';

/**
 * The landing screen: what the list adds up to, and a way into the rows behind
 * every figure on it.
 *
 * Every number here is derived from the same fetch the records screen uses, so
 * a card and the list it opens can never disagree. There is no server-side
 * aggregate to call — the SharePoint list is the only source this app has.
 */

function countBy(items, pick) {
  const counts = new Map();
  for (const item of items) {
    const key = pick(item);
    if (!key) continue;
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  return [...counts.entries()].map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value);
}

export default function DashboardPage() {
  const navigate = useNavigate();
  const { items, loading, error, loadedAt, reload } = useRequests();
  const { markContentReady } = useSession();

  useEffect(() => {
    document.title = 'PMW IT — Dashboard';
  }, []);

  // The screen someone lands on after signing in, so it is the screen the
  // entrance animation waits for: it holds until the figures behind it are real
  // rather than fading out over a page of skeletons. A failed load counts as
  // ready too — the error belongs on the dashboard, not under a veil.
  useEffect(() => {
    if (!loading) markContentReady();
  }, [loading, markContentReady]);

  const stats = useMemo(() => {
    const now = new Date();
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

    const typeOf = (item) => String(item.Request_x0020_Type || '').toLowerCase();

    const monthly = [];
    for (let back = 5; back >= 0; back -= 1) {
      const month = new Date(now.getFullYear(), now.getMonth() - back, 1);
      const next = new Date(now.getFullYear(), now.getMonth() - back + 1, 1);
      monthly.push({
        label: month.toLocaleString(undefined, { month: 'short' }),
        value: items.filter((item) => {
          const d = requestDate(item);
          return d && d >= month && d < next;
        }).length,
      });
    }

    const equipment = new Map();
    for (const item of items) {
      for (const piece of toChoiceArray(item.Equipment_x0020_Items)) {
        equipment.set(piece, (equipment.get(piece) || 0) + 1);
      }
    }

    return {
      total: items.length,
      onboarding: items.filter((item) => typeOf(item) === 'onboarding').length,
      offboarding: items.filter((item) => typeOf(item) === 'offboarding').length,
      thisMonth: items.filter((item) => {
        const d = requestDate(item);
        return d && d >= startOfMonth;
      }).length,
      lastWeek: items.filter((item) => {
        const d = requestDate(item);
        return d && d >= weekAgo;
      }).length,
      withEquipment: items.filter((item) => toChoiceArray(item.Equipment_x0020_Items).length > 0).length,
      monthly,
      byEntity: countBy(items, (item) => item.Entity),
      byDepartment: countBy(items, (item) => item.Department).slice(0, 6),
      byEquipment: [...equipment.entries()]
        .map(([label, value]) => ({ label, value }))
        .sort((a, b) => b.value - a.value),
      recent: [...items]
        .sort((a, b) => (requestDate(b)?.getTime() || 0) - (requestDate(a)?.getTime() || 0))
        .slice(0, 6),
    };
  }, [items]);

  const snapshotAt = loadedAt
    ? loadedAt.toLocaleTimeString(undefined, { hour: '2-digit', minute: '2-digit' })
    : null;

  const CARDS = [
    {
      key: 'total',
      label: 'Total requests',
      icon: ClipboardList,
      color: 'var(--it-brand)',
      value: stats.total,
      to: '/requests',
    },
    {
      key: 'onboarding',
      label: 'Onboarding',
      icon: UserPlus,
      color: 'var(--it-good)',
      value: stats.onboarding,
      to: '/requests?type=Onboarding',
    },
    {
      key: 'offboarding',
      label: 'Offboarding',
      icon: UserMinus,
      color: 'var(--it-danger)',
      value: stats.offboarding,
      to: '/requests?type=Offboarding',
    },
    {
      key: 'month',
      label: 'This month',
      icon: Calendar,
      color: 'var(--it-brand-mid)',
      value: stats.thisMonth,
      to: '/requests?range=month',
    },
    {
      key: 'week',
      label: 'Last 7 days',
      icon: Clock,
      color: 'var(--it-accent)',
      value: stats.lastWeek,
      to: '/requests?range=week',
    },
    {
      key: 'equipment',
      label: 'With equipment',
      icon: Laptop,
      color: 'var(--it-brand)',
      value: stats.withEquipment,
      to: '/requests?equipment=yes',
    },
  ];

  return (
    <AppShell
      title="Dashboard"
      subtitle={
        snapshotAt
          ? `Read at ${snapshotAt} · select any card for the requests behind it`
          : 'Loading the request list…'
      }
      actions={
        <>
          <Button variant="ghost" icon={RefreshCw} onClick={reload} disabled={loading}>
            {loading ? 'Refreshing…' : 'Refresh'}
          </Button>
          <Button icon={Plus} onClick={() => navigate('/it-boarding-form')}>
            New request
          </Button>
        </>
      }
    >
      {error && <ErrorBanner message={error} onRetry={reload} />}

      <div className="stat-grid">
        {CARDS.map((card) => (
          <StatCard
            key={card.key}
            icon={card.icon}
            label={card.label}
            color={card.color}
            value={card.value}
            loading={loading}
            onClick={() => navigate(card.to)}
          />
        ))}
      </div>

      <div className="chart-grid">
        <ColumnChart
          title="Requests per month"
          blurb="The last six months, by join or last working date."
          columns={stats.monthly}
        />
        <BarChart
          title="By entity"
          blurb="Which company each request was raised for."
          rows={stats.byEntity}
          emptyText="No entity recorded on any request yet."
          onSelect={(row) => navigate(`/requests?entity=${encodeURIComponent(row.label)}`)}
        />
        <BarChart
          title="By department"
          blurb="The six departments raising the most requests."
          rows={stats.byDepartment}
          emptyText="No department recorded on any request yet."
          onSelect={(row) => navigate(`/requests?department=${encodeURIComponent(row.label)}`)}
        />
        <BarChart
          title="Equipment requested"
          blurb="How often each item appears across all requests."
          rows={stats.byEquipment}
          emptyText="No equipment has been requested yet."
        />

        <Card className="chart-card">
          <div className="chart-head">
            <h3>Latest requests</h3>
            <p>The six most recent, newest first. Select one to open it.</p>
          </div>
          {loading ? (
            <EmptyState>Loading…</EmptyState>
          ) : stats.recent.length === 0 ? (
            <EmptyState>Nothing has been submitted yet.</EmptyState>
          ) : (
            stats.recent.map((item) => (
              <button
                type="button"
                className="activity-row"
                key={item.ID}
                onClick={() => navigate(`/it-boarding-form?edit=${item.ID}`)}
              >
                <div className="activity-main">
                  <div className="activity-title">{item.Title || 'Untitled request'}</div>
                  <div className="activity-meta">
                    {[item.Position, item.Department, item.Entity].filter(Boolean).join(' · ') || 'No details'}
                    {' · '}
                    {formatDate(requestDate(item))}
                  </div>
                </div>
                <RequestTypeBadge type={item.Request_x0020_Type} />
                <ChevronRight size={15} style={{ color: 'var(--it-ink-soft)', flexShrink: 0 }} />
              </button>
            ))
          )}
        </Card>
      </div>
    </AppShell>
  );
}

import { useEffect, useMemo, useState } from 'react';
import { useNavigate, useSearchParams } from 'react-router-dom';
import AppShell from '../components/AppShell';
import { initialsOf } from '../utils/initials';
import Button from '../components/ui/Button';
import { EmptyState, ErrorBanner } from '../components/ui/Surfaces';
import { RequestTypeBadge } from '../components/ui/Badges';
import { Search, Filter, Plus, RefreshCw, Pencil, Copy, Inbox } from '../components/ui/Icons';
import { useRequests, requestDate, formatDate, toChoiceArray } from '../hooks/useRequests';

/**
 * The records screen: every request, and the ways of cutting it down.
 *
 * The filters live in the query string rather than in component state alone,
 * which is what lets a dashboard card open its own slice of this list — and
 * what makes "Copy link" worth having, since the link carries the view.
 */

const SORTS = [
  { value: 'newest', label: 'Newest first' },
  { value: 'oldest', label: 'Oldest first' },
  { value: 'name', label: 'Name A–Z' },
  { value: 'position', label: 'Position' },
  { value: 'entity', label: 'Entity' },
];

export default function ListPage() {
  const navigate = useNavigate();
  const [searchParams, setSearchParams] = useSearchParams();
  const { items, choices, loading, error, reload } = useRequests();

  useEffect(() => {
    document.title = 'PMW IT — Requests';
  }, []);

  // Read once, on mount: after that this screen owns the values and writes them
  // back out. Reading them continuously would fight the inputs.
  const [query, setQuery] = useState(() => searchParams.get('q') || '');
  const [sortBy, setSortBy] = useState(() => searchParams.get('sort') || 'newest');
  const [entity, setEntity] = useState(() => searchParams.get('entity') || '');
  const [type, setType] = useState(() => searchParams.get('type') || '');
  const [department, setDepartment] = useState(() => searchParams.get('department') || '');
  const [dateFrom, setDateFrom] = useState(() => searchParams.get('from') || '');
  const [dateTo, setDateTo] = useState(() => searchParams.get('to') || '');
  const [range, setRange] = useState(() => searchParams.get('range') || '');
  const [equipmentOnly, setEquipmentOnly] = useState(() => searchParams.get('equipment') === 'yes');
  const [showFilters, setShowFilters] = useState(false);
  const [copied, setCopied] = useState(false);

  // Push the current view back into the URL so it can be linked to. `replace`
  // keeps a filter session out of the back button's history.
  useEffect(() => {
    const next = new URLSearchParams();
    if (query) next.set('q', query);
    if (sortBy && sortBy !== 'newest') next.set('sort', sortBy);
    if (entity) next.set('entity', entity);
    if (type) next.set('type', type);
    if (department) next.set('department', department);
    if (dateFrom) next.set('from', dateFrom);
    if (dateTo) next.set('to', dateTo);
    if (range) next.set('range', range);
    if (equipmentOnly) next.set('equipment', 'yes');
    setSearchParams(next, { replace: true });
  }, [query, sortBy, entity, type, department, dateFrom, dateTo, range, equipmentOnly, setSearchParams]);

  const clearAll = () => {
    setEntity('');
    setType('');
    setDepartment('');
    setDateFrom('');
    setDateTo('');
    setRange('');
    setEquipmentOnly(false);
  };

  const filtered = useMemo(() => {
    let result = [...items];

    if (query) {
      const q = query.toLowerCase();
      result = result.filter((item) =>
        [item.Title, item.Position, item.Entity, item.Calling_x0020_Name, item.Department, item.Employee_x0020_ID]
          .some((field) => String(field || '').toLowerCase().includes(q))
      );
    }

    if (entity) result = result.filter((item) => item.Entity === entity);
    if (type) result = result.filter((item) => item.Request_x0020_Type === type);
    if (department) result = result.filter((item) => item.Department === department);
    if (equipmentOnly) result = result.filter((item) => toChoiceArray(item.Equipment_x0020_Items).length > 0);

    if (range) {
      const now = new Date();
      const floor =
        range === 'week'
          ? new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000)
          : new Date(now.getFullYear(), now.getMonth(), 1);
      result = result.filter((item) => {
        const d = requestDate(item);
        return d && d >= floor;
      });
    }

    if (dateFrom) {
      const from = new Date(dateFrom);
      result = result.filter((item) => {
        const d = requestDate(item);
        return d && d >= from;
      });
    }

    if (dateTo) {
      const to = new Date(dateTo);
      result = result.filter((item) => {
        const d = requestDate(item);
        return d && d <= to;
      });
    }

    result.sort((a, b) => {
      switch (sortBy) {
        case 'oldest':
          return (requestDate(a)?.getTime() || 0) - (requestDate(b)?.getTime() || 0);
        case 'name':
          return String(a.Title || '').localeCompare(String(b.Title || ''));
        case 'position':
          return String(a.Position || '').localeCompare(String(b.Position || ''));
        case 'entity':
          return String(a.Entity || '').localeCompare(String(b.Entity || ''));
        case 'newest':
        default:
          return (requestDate(b)?.getTime() || 0) - (requestDate(a)?.getTime() || 0);
      }
    });

    return result;
  }, [items, query, sortBy, entity, type, department, dateFrom, dateTo, range, equipmentOnly]);

  const activeFilters = [entity, type, department, dateFrom, dateTo, range, equipmentOnly ? 'yes' : ''].filter(
    Boolean
  ).length;

  const copyLink = async () => {
    try {
      await navigator.clipboard.writeText(window.location.href);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      // Clipboard access can be refused (an insecure origin, a locked-down
      // browser). Nothing here is worth an error dialog over.
      setCopied(false);
    }
  };

  return (
    <AppShell
      title="Requests"
      subtitle={
        loading ? 'Loading the request list…' : `Showing ${filtered.length} of ${items.length} requests`
      }
      search={{ value: query, onChange: setQuery, placeholder: 'Search name, position, entity…' }}
      actions={
        <>
          <Button variant="ghost" icon={Copy} onClick={copyLink}>
            {copied ? 'Link copied' : 'Copy link'}
          </Button>
          <Button variant="ghost" icon={RefreshCw} onClick={reload} disabled={loading}>
            Refresh
          </Button>
          <Button icon={Plus} onClick={() => navigate('/it-boarding-form')}>
            New request
          </Button>
        </>
      }
    >
      {error && <ErrorBanner message={error} onRetry={reload} />}

      <div className="toolbar">
        {/* The bar carries this control from 640px up; below that it is hidden,
            and this is the one on screen. */}
        <div className="search-box only-narrow-flex">
          <Search size={18} />
          <input
            type="text"
            placeholder="Search name, position, entity…"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
          {query && (
            <button type="button" className="clear-btn" onClick={() => setQuery('')} aria-label="Clear search">
              ×
            </button>
          )}
        </div>

        <div className="toolbar-actions">
          <button
            type="button"
            className={`filter-btn ${showFilters ? 'active' : ''}`}
            onClick={() => setShowFilters((v) => !v)}
            aria-expanded={showFilters}
          >
            <Filter size={18} />
            Filters
            {activeFilters > 0 && <span className="filter-badge">{activeFilters}</span>}
          </button>

          <select
            value={sortBy}
            onChange={(e) => setSortBy(e.target.value)}
            className="sort-select"
            aria-label="Sort requests"
          >
            {SORTS.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>
      </div>

      {showFilters && (
        <div className="filter-panel">
          <div className="filter-group">
            <label htmlFor="filter-entity">Entity</label>
            <select id="filter-entity" value={entity} onChange={(e) => setEntity(e.target.value)}>
              <option value="">All entities</option>
              {(choices.Entity || []).map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </div>

          <div className="filter-group">
            <label htmlFor="filter-type">Request type</label>
            <select id="filter-type" value={type} onChange={(e) => setType(e.target.value)}>
              <option value="">All types</option>
              {(choices.Request_x0020_Type || []).map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </div>

          <div className="filter-group">
            <label htmlFor="filter-department">Department</label>
            <select id="filter-department" value={department} onChange={(e) => setDepartment(e.target.value)}>
              <option value="">All departments</option>
              {(choices.Department || []).map((value) => (
                <option key={value} value={value}>
                  {value}
                </option>
              ))}
            </select>
          </div>

          <div className="filter-group">
            <label htmlFor="filter-range">Period</label>
            <select id="filter-range" value={range} onChange={(e) => setRange(e.target.value)}>
              <option value="">Any time</option>
              <option value="week">Last 7 days</option>
              <option value="month">This month</option>
            </select>
          </div>

          <div className="filter-group">
            <label htmlFor="filter-from">Date from</label>
            <input
              id="filter-from"
              type="date"
              value={dateFrom}
              onChange={(e) => setDateFrom(e.target.value)}
            />
          </div>

          <div className="filter-group">
            <label htmlFor="filter-to">Date to</label>
            <input id="filter-to" type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)} />
          </div>

          <button type="button" className="clear-filters-btn" onClick={clearAll}>
            Clear all
          </button>
        </div>
      )}

      {loading ? (
        <div className="loading-card">
          <div className="spinner" />
          <p>Loading requests…</p>
        </div>
      ) : filtered.length === 0 ? (
        <div className="ui-card">
          <EmptyState>
            <Inbox size={40} style={{ marginBottom: 10 }} />
            <div style={{ fontSize: 15, fontWeight: 700, color: 'var(--it-ink)', marginBottom: 4 }}>
              {items.length === 0 ? 'No requests yet' : 'Nothing matches this view'}
            </div>
            <p style={{ margin: '0 0 14px' }}>
              {items.length === 0
                ? 'Raise the first onboarding or offboarding request to get started.'
                : 'Try a different search, or clear the filters.'}
            </p>
            {items.length === 0 ? (
              <Button icon={Plus} onClick={() => navigate('/it-boarding-form')}>
                New request
              </Button>
            ) : (
              <Button
                variant="ghost"
                onClick={() => {
                  setQuery('');
                  clearAll();
                }}
              >
                Clear search and filters
              </Button>
            )}
          </EmptyState>
        </div>
      ) : (
        <div className="list-table">
          <table>
            <thead>
              <tr>
                <th>Employee</th>
                <th>Position</th>
                <th>Entity</th>
                <th>Department</th>
                <th>Request type</th>
                <th>Date</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((item) => (
                <tr key={item.ID}>
                  <td>
                    <div className="employee-cell">
                      <div className="employee-avatar">{initialsOf(item.Title)}</div>
                      <div className="employee-info">
                        <span className="employee-name">{item.Title || '-'}</span>
                        <span className="employee-callname">{item.Calling_x0020_Name || ''}</span>
                      </div>
                    </div>
                  </td>
                  <td>{item.Position || '-'}</td>
                  <td>{item.Entity || '-'}</td>
                  <td>{item.Department || '-'}</td>
                  <td>
                    <RequestTypeBadge type={item.Request_x0020_Type} />
                  </td>
                  <td>{formatDate(requestDate(item))}</td>
                  <td>
                    <div className="action-buttons">
                      <button
                        type="button"
                        className="action-btn edit-btn"
                        onClick={() => navigate(`/it-boarding-form?edit=${item.ID}`)}
                        title="Edit request"
                        aria-label={`Edit the request for ${item.Title || 'this employee'}`}
                      >
                        <Pencil size={16} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </AppShell>
  );
}

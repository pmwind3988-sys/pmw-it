import React, { useState, useEffect, useRef } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { useTheme } from '../context/ThemeContext';
import { fetchAllListItems, fetchAllColumnChoices, updateListItem } from '../services/sharePointService';
import { sharePointRequest } from '../authConfig';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL ||
  'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const CHOICE_COLUMNS = ['Entity', 'Equipment_x0020_Items', 'Software_x0020_Licenses', 'Request_x0020_Type'];

export default function ListPage() {
  const { instance } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const { isDarkMode, toggleTheme } = useTheme();

  useEffect(() => {
    document.title = 'IT REQUEST FORM';
  }, []);

  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [spChoices, setSpChoices] = useState({});
  const [editingItem, setEditingItem] = useState(null);
  const [showSharePanel, setShowSharePanel] = useState(false);
  const sharePanelRef = useRef(null);
  
  // Search, sort, filter states
  const [searchQuery, setSearchQuery] = useState('');
  const [sortBy, setSortBy] = useState('newest');
  const [filterEntity, setFilterEntity] = useState('');
  const [filterType, setFilterType] = useState('');
  const [filterDateFrom, setFilterDateFrom] = useState('');
  const [filterDateTo, setFilterDateTo] = useState('');
  const [showFilters, setShowFilters] = useState(false);

  const getToken = async () => {
    const account = instance.getActiveAccount();
    if (!account) throw new Error('No signed-in account');
    try {
      return await instance.acquireTokenSilent({ ...sharePointRequest, account });
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        return await instance.acquireTokenPopup({ ...sharePointRequest, account });
      }
      throw e;
    }
  };

  useEffect(() => {
    if (!isAuthenticated) return;
    let cancelled = false;

    async function loadData() {
      setLoading(true);
      setError('');
      try {
        const tokenRes = await getToken();
        const [itemsData, choices] = await Promise.all([
          fetchAllListItems(SHAREPOINT_SITE_URL, tokenRes.accessToken),
          fetchAllColumnChoices(SHAREPOINT_SITE_URL, tokenRes.accessToken, CHOICE_COLUMNS),
        ]);
        if (!cancelled) {
          setItems(itemsData);
          setSpChoices(choices);
        }
      } catch (err) {
        if (!cancelled) setError(err.message || 'Failed to load data');
      } finally {
        if (!cancelled) setLoading(false);
      }
    }
    loadData();
    return () => { cancelled = true; };
  }, [isAuthenticated]);

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (sharePanelRef.current && !sharePanelRef.current.contains(event.target)) {
        setShowSharePanel(false);
      }
    };
    if (showSharePanel) {
      document.addEventListener('mousedown', handleClickOutside);
      return () => document.removeEventListener('mousedown', handleClickOutside);
    }
  }, [showSharePanel]);

  const logout = () => instance.logoutRedirect({ postLogoutRedirectUri: import.meta.env.VITE_REDIRECT_URI || 'http://localhost:5173' });

  const getInitials = (name) => {
    if (!name) return 'U';
    const parts = name.split(' ');
    return parts.length >= 2 ? (parts[0][0] + parts[1][0]).toUpperCase() : name.substring(0, 2).toUpperCase();
  };

  const formatDate = (dateStr) => {
    if (!dateStr) return '-';
    const d = new Date(dateStr);
    return isNaN(d.getTime()) ? '-' : d.toLocaleDateString();
  };

  const formatChoices = (arr) => Array.isArray(arr) ? arr.join(', ') : '-';

  // Filter and sort items
  const filteredItems = React.useMemo(() => {
    let result = [...items];
    
    // Search filter
    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      result = result.filter(item => 
        (item.Title || '').toLowerCase().includes(q) ||
        (item.Position || '').toLowerCase().includes(q) ||
        (item.Entity || '').toLowerCase().includes(q) ||
        (item.Calling_x0020_Name || '').toLowerCase().includes(q)
      );
    }
    
    // Entity filter
    if (filterEntity) {
      result = result.filter(item => item.Entity === filterEntity);
    }
    
    // Request Type filter
    if (filterType) {
      result = result.filter(item => item.Request_x0020_Type === filterType);
    }
    
    // Date from filter
    if (filterDateFrom) {
      const fromDate = new Date(filterDateFrom);
      result = result.filter(item => {
        if (!item.Join_x0020__x002f__x0020_Last_x0) return false;
        const itemDate = new Date(item.Join_x0020__x002f__x0020_Last_x0);
        return itemDate >= fromDate;
      });
    }
    
    // Date to filter
    if (filterDateTo) {
      const toDate = new Date(filterDateTo);
      result = result.filter(item => {
        if (!item.Join_x0020__x002f__x0020_Last_x0) return false;
        const itemDate = new Date(item.Join_x0020__x002f__x0020_Last_x0);
        return itemDate <= toDate;
      });
    }
    
    // Sort
    result.sort((a, b) => {
      switch (sortBy) {
        case 'newest':
          return new Date(b.Join_x0020__x002f__x0020_Last_x0) - new Date(a.Join_x0020__x002f__x0020_Last_x0);
        case 'oldest':
          return new Date(a.Join_x0020__x002f__x0020_Last_x0) - new Date(b.Join_x0020__x002f__x0020_Last_x0);
        case 'name':
          return (a.Title || '').localeCompare(b.Title || '');
        case 'position':
          return (a.Position || '').localeCompare(b.Position || '');
        case 'entity':
          return (a.Entity || '').localeCompare(b.Entity || '');
        default:
          return 0;
      }
    });
    
    return result;
  }, [items, searchQuery, sortBy, filterEntity, filterType, filterDateFrom, filterDateTo]);

  // Active filters count
  const activeFiltersCount = [filterEntity, filterType, filterDateFrom, filterDateTo].filter(Boolean).length;

  if (!isAuthenticated) {
    return (
      <div className="login-required">
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <circle cx="12" cy="12" r="10" />
          <path d="M12 16v-4M12 8h.01" />
        </svg>
        <h2>Sign in Required</h2>
        <p>Please log in to access this page.</p>
        <button className="ms-button" onClick={() => window.location.href = '/login'}>Sign In</button>
      </div>
    );
  }

  const account = instance.getActiveAccount();

  return (
    <div className="form-page">
      <div className="auth-banner">
        <div className="auth-banner-left">
          {account && (
            <>
              <div className="user-avatar">{getInitials(account.name)}</div>
              <span className="user-name">{account.name}</span>
            </>
          )}
        </div>
        <div className="auth-banner-right">
          <button className="icon-btn" onClick={() => window.location.href = '/it-boarding-form'} title="Add New">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="12" cy="12" r="10" /><line x1="12" y1="8" x2="12" y2="16" /><line x1="8" y1="12" x2="16" y2="12" />
            </svg>
          </button>
          <button className="icon-btn" onClick={() => setShowSharePanel((v) => !v)} title="Share">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="18" cy="5" r="3" /><circle cx="6" cy="12" r="3" /><circle cx="18" cy="19" r="3" />
              <line x1="8.59" y1="13.51" x2="15.42" y2="17.49" /><line x1="15.41" y1="6.51" x2="8.59" y2="10.49" />
            </svg>
          </button>
          <button className="icon-btn" onClick={toggleTheme} title={isDarkMode ? 'Light Mode' : 'Dark Mode'}>
            {isDarkMode ? (
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="5" /><line x1="12" y1="1" x2="12" y2="3" /><line x1="12" y1="21" x2="12" y2="23" />
                <line x1="4.22" y1="4.22" x2="5.64" y2="5.64" /><line x1="18.36" y1="18.36" x2="19.78" y2="19.78" />
                <line x1="1" y1="12" x2="3" y2="12" /><line x1="21" y1="12" x2="23" y2="12" />
                <line x1="4.22" y1="19.78" x2="5.64" y2="18.36" /><line x1="18.36" y1="5.64" x2="19.78" y2="4.22" />
              </svg>
            ) : (
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" />
              </svg>
            )}
          </button>
          <button className="icon-btn" onClick={logout} title="Logout">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" />
              <polyline points="16 17 21 12 16 7" /><line x1="21" y1="12" x2="9" y2="12" />
            </svg>
          </button>
        </div>
      </div>

      {showSharePanel && (
        <div className="share-panel" ref={sharePanelRef}>
          <div className="share-panel-item" onClick={async () => { await navigator.clipboard.writeText(window.location.href); setShowSharePanel(false); }}>
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <rect x="9" y="9" width="13" height="13" rx="2" ry="2" /><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
            </svg>
            <span>Copy Link</span>
          </div>
        </div>
      )}

      <div className="form-content">
        <div className="form-header">
          <h1>IT Request Form</h1>
          <p>View and manage submitted requests</p>
        </div>

        {/* Search and Filter Toolbar */}
        <div className="toolbar">
          <div className="search-box">
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" />
            </svg>
            <input 
              type="text" 
              placeholder="Search by name, position, entity..." 
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
            />
            {searchQuery && (
              <button className="clear-btn" onClick={() => setSearchQuery('')}>×</button>
            )}
          </div>
          
          <div className="toolbar-actions">
            <button className={`filter-btn ${showFilters ? 'active' : ''}`} onClick={() => setShowFilters(!showFilters)}>
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" />
              </svg>
              Filters
              {activeFiltersCount > 0 && <span className="filter-badge">{activeFiltersCount}</span>}
            </button>
            
            <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} className="sort-select">
              <option value="newest">Newest First</option>
              <option value="oldest">Oldest First</option>
              <option value="name">Name A-Z</option>
              <option value="position">Position</option>
              <option value="entity">Entity</option>
            </select>
          </div>
        </div>

        {/* Filter Panel */}
        {showFilters && (
          <div className="filter-panel">
            <div className="filter-group">
              <label>Entity</label>
              <select value={filterEntity} onChange={(e) => setFilterEntity(e.target.value)}>
                <option value="">All Entities</option>
                {(spChoices['Entity'] || []).map(e => (
                  <option key={e} value={e}>{e}</option>
                ))}
              </select>
            </div>
            
            <div className="filter-group">
              <label>Request Type</label>
              <select value={filterType} onChange={(e) => setFilterType(e.target.value)}>
                <option value="">All Types</option>
                {(spChoices['Request_x0020_Type'] || []).map(t => (
                  <option key={t} value={t}>{t}</option>
                ))}
              </select>
            </div>
            
            <div className="filter-group">
              <label>Date From</label>
              <input type="date" value={filterDateFrom} onChange={(e) => setFilterDateFrom(e.target.value)} />
            </div>
            
            <div className="filter-group">
              <label>Date To</label>
              <input type="date" value={filterDateTo} onChange={(e) => setFilterDateTo(e.target.value)} />
            </div>
            
            <button className="clear-filters-btn" onClick={() => {
              setFilterEntity('');
              setFilterType('');
              setFilterDateFrom('');
              setFilterDateTo('');
            }}>
              Clear All
            </button>
          </div>
        )}

        {/* Results Count */}
        <div className="results-info">
          Showing {filteredItems.length} of {items.length} requests
        </div>

        {loading ? (
          <div className="loading-card">
            <div className="spinner"></div>
            <p>Loading requests…</p>
          </div>
        ) : error ? (
          <div className="error-screen">
            <p className="error-message">{error}</p>
            <button className="ms-button" onClick={() => window.location.reload()}>Retry</button>
          </div>
        ) : filteredItems.length === 0 ? (
          <div className="empty-list">
            {items.length === 0 ? (
              <>
                <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
                  <polyline points="14 2 14 8 20 8" /><line x1="12" y1="18" x2="12" y2="12" /><line x1="9" y1="15" x2="15" y2="15" />
                </svg>
                <h2>No Requests Yet</h2>
                <p>Click the + button to add a new request</p>
                <button className="ms-button" onClick={() => window.location.href = '/it-boarding-form'}>Add Request</button>
              </>
            ) : (
              <>
                <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" />
                </svg>
                <h2>No Results Found</h2>
                <p>Try adjusting your search or filters</p>
                <button className="ms-button" onClick={() => {
                  setSearchQuery('');
                  setFilterEntity('');
                  setFilterType('');
                  setFilterDateFrom('');
                  setFilterDateTo('');
                }}>Clear All</button>
              </>
            )}
          </div>
        ) : (
          <div className="list-table">
            <table>
              <thead>
                <tr>
                  <th>Employee</th>
                  <th>Position</th>
                  <th>Entity</th>
                  <th>Request Type</th>
                  <th>Date</th>
                  <th>Actions</th>
                </tr>
              </thead>
              <tbody>
                {filteredItems.map((item) => (
                  <tr key={item.ID}>
                    <td>
                      <div className="employee-cell">
                        <div className="employee-avatar">{getInitials(item.Title)}</div>
                        <div className="employee-info">
                          <span className="employee-name">{item.Title || '-'}</span>
                          <span className="employee-callname">{item.Calling_x0020_Name || ''}</span>
                        </div>
                      </div>
                    </td>
                    <td>{item.Position || '-'}</td>
                    <td>{item.Entity || '-'}</td>
                    <td><span className={`badge badge-${item.Request_x0020_Type?.toLowerCase()}`}>{item.Request_x0020_Type || '-'}</span></td>
                    <td>{formatDate(item.Join_x0020__x002f__x0020_Last_x0)}</td>
                    <td>
                      <div className="action-buttons">
                        <button className="action-btn view-btn" onClick={() => window.location.href = `/it-boarding-form?edit=${item.ID}`} title="View">
                          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" /><circle cx="12" cy="12" r="3" />
                          </svg>
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
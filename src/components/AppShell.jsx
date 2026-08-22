import { useEffect, useState } from 'react';
import { Link, useLocation, useNavigate } from 'react-router-dom';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';
import { useTheme } from '../context/ThemeContext';
import { useSession } from '../hooks/useSession';
import { initialsOf } from '../utils/initials';
import Logo from './Logo';
import { AccountBadge } from './ui/Badges';
import Button from './ui/Button';
import {
  LayoutDashboard,
  ClipboardList,
  FilePlus,
  CheckSquare,
  Menu,
  X,
  Search,
  LogOut,
  Sun,
  Moon,
  ShieldCheck,
  Laptop,
  BarChart3,
} from './ui/Icons';

/**
 * Two layouts from one tree, switched at 1024px:
 *
 *   < lg  — the brand column is an off-canvas drawer behind a hamburger. As a
 *           permanent column it left a 360px phone with about 130px for the
 *           actual screen, which made every page unusable.
 *   >= lg — the same column is sticky, exactly as the SI shell has it.
 *
 * The switch is pure CSS (a transform plus the 1024px overrides in shell.css)
 * rather than a JS width check, so there is no flash of the wrong layout on
 * first paint.
 */

const NAV_ITEMS = [
  { to: '/dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { to: '/requests', label: 'Requests', icon: ClipboardList },
  { to: '/it-boarding-form', label: 'New request', icon: FilePlus },
  { to: '/asset-checklist', label: 'Asset checklist', icon: CheckSquare },
  { to: '/devices', label: 'Device list', icon: Laptop },
  { to: '/data-studio', label: 'Data Studio', icon: BarChart3 },
];

export default function AppShell({ title, subtitle, actions, search, children }) {
  const { instance, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const { isDarkMode, toggleTheme } = useTheme();
  // Signing out belongs to the session guard: it has to forget who was here
  // before the redirect, or the automatic sign-in reads the sign-out as a
  // timeout and puts the user straight back in.
  const { signOut, recovering } = useSession();
  const location = useLocation();
  const navigate = useNavigate();
  const [navOpen, setNavOpen] = useState(false);
  const [headerQuery, setHeaderQuery] = useState('');

  // Navigating from the drawer must close it, or the page just asked for stays
  // hidden behind it. Done on the links themselves rather than in an effect on
  // the path: they are the only things inside the drawer that navigate, and an
  // effect that calls setState on every route change costs a second render of
  // the whole shell for the one case where the drawer was open.
  const closeNav = () => setNavOpen(false);

  // While the drawer is over the page, the page behind it must not scroll — on
  // a phone a scrolling backdrop reads as the drawer itself failing to scroll.
  // Escape closes it for anyone on a keyboard.
  useEffect(() => {
    if (!navOpen) return undefined;
    const onKey = (e) => {
      if (e.key === 'Escape') setNavOpen(false);
    };
    document.addEventListener('keydown', onKey);
    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    return () => {
      document.removeEventListener('keydown', onKey);
      document.body.style.overflow = previousOverflow;
    };
  }, [navOpen]);

  const account = instance.getActiveAccount();

  // The shell is the one auth gate: every screen inside it needs an account, so
  // each page used to carry its own copy of this block.
  if (inProgress !== InteractionStatus.None && !isAuthenticated) {
    return (
      <div className="shell-gate">
        <div className="spinner" />
        <p>Checking your sign-in…</p>
      </div>
    );
  }

  // A timed-out session is being put right — the guard's dialog is over this,
  // saying so. Announcing "sign in required" underneath it would contradict it,
  // and the button would race the recovery.
  if (!isAuthenticated && recovering) {
    return (
      <div className="shell-gate">
        <div className="spinner" />
        <p>Restoring your session…</p>
      </div>
    );
  }

  if (!isAuthenticated) {
    return (
      <div className="shell-gate">
        <ShieldCheck size={40} />
        <h2>Sign in required</h2>
        <p>Sign in with your Microsoft 365 work account to reach this page.</p>
        <Button onClick={() => navigate('/login')}>Go to sign in</Button>
      </div>
    );
  }

  const isActive = (to) => location.pathname === to || location.pathname.startsWith(`${to}/`);

  // Without a `search` prop the bar still carries a box, and it takes you to the
  // records with the query applied — the same box on every screen, one job.
  const searchValue = search ? search.value : headerQuery;
  const onSearchChange = (value) => {
    if (search) search.onChange(value);
    else setHeaderQuery(value);
  };
  const onSearchSubmit = (e) => {
    e.preventDefault();
    if (!search && headerQuery.trim()) navigate(`/requests?q=${encodeURIComponent(headerQuery.trim())}`);
  };

  return (
    <div className="shell">
      {navOpen && (
        <button
          type="button"
          className="shell-backdrop"
          onClick={closeNav}
          aria-label="Close navigation"
        />
      )}

      <aside
        id="app-nav"
        aria-label="Main navigation"
        className={`shell-nav it-brand-surface${navOpen ? ' open' : ''}`}
      >
        <div className="shell-brand">
          <Link to="/dashboard" className="shell-brand-link" onClick={closeNav}>
            <span className="shell-logo-chip">
              <Logo size={26} />
            </span>
            <span>
              <span className="shell-brand-name">PMW IT</span>
              <span className="shell-brand-sub">Service Portal</span>
            </span>
          </Link>
          <button
            type="button"
            className="shell-navclose"
            onClick={closeNav}
            aria-label="Close navigation"
          >
            <X size={20} />
          </button>
        </div>

        <nav className="shell-navlist">
          {NAV_ITEMS.map((item) => (
            <Link
              key={item.to}
              to={item.to}
              className={`shell-navitem${isActive(item.to) ? ' active' : ''}`}
              aria-current={isActive(item.to) ? 'page' : undefined}
              onClick={closeNav}
            >
              <item.icon size={16} />
              {item.label}
            </Link>
          ))}
        </nav>

        <div className="shell-navfoot">
          <div className="shell-navuser">
            <div className="shell-avatar">{initialsOf(account?.name)}</div>
            <div className="shell-navuser-text">
              <div className="shell-navuser-name">{account?.name || 'Signed in'}</div>
              <div className="shell-navuser-mail">{account?.username || 'Microsoft 365'}</div>
            </div>
          </div>
          <button type="button" className="shell-signout" onClick={signOut}>
            <LogOut size={13} /> Sign out
          </button>
        </div>
      </aside>

      <div className="shell-main">
        <header className="shell-header">
          <div className="shell-headerrow">
            <button
              type="button"
              className="shell-hamburger"
              onClick={() => setNavOpen(true)}
              aria-label="Open navigation"
              aria-controls="app-nav"
              aria-expanded={navOpen}
            >
              <Menu size={22} />
            </button>

            {/* The column's brand mark is behind the drawer on a phone, so the
                bar carries one of its own. */}
            <Link to="/dashboard" className="shell-headerbrand" aria-label="PMW IT — Service Portal">
              <Logo size={24} />
              <span className="hide-below-xs">PMW IT</span>
            </Link>

            <form className="shell-search" onSubmit={onSearchSubmit} role="search">
              <Search size={15} />
              <input
                value={searchValue}
                onChange={(e) => onSearchChange(e.target.value)}
                placeholder={search?.placeholder || 'Search requests…'}
                aria-label="Search requests"
              />
            </form>

            <div className="shell-headeractions">
              <button
                type="button"
                className="shell-iconbtn"
                onClick={toggleTheme}
                title={isDarkMode ? 'Light mode' : 'Dark mode'}
                aria-label={isDarkMode ? 'Switch to light mode' : 'Switch to dark mode'}
              >
                {isDarkMode ? <Sun size={17} /> : <Moon size={17} />}
              </button>
              <AccountBadge name={account?.name} compact />
            </div>
          </div>
        </header>

        {/* Keyed on the path so the entrance replays on every navigation rather
            than only on first mount — which is the point of it: it marks that
            the content changed, on a phone where there is no other cue. */}
        <main key={location.pathname} className="shell-body rise">
          {(title || actions) && (
            <div className="shell-pagehead">
              <div>
                {title && <h1>{title}</h1>}
                {subtitle && <p>{subtitle}</p>}
              </div>
              {actions && <div className="shell-pagehead-actions">{actions}</div>}
            </div>
          )}
          {children}
        </main>
      </div>
    </div>
  );
}

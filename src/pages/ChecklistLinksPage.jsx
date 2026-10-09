import { Fragment, useCallback, useEffect, useState } from 'react';
import { useNavigate, useSearchParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import Spinner from '../components/ui/Spinner';
import { Card, ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import {
  Copy, Check, Plus, X, RefreshCw, Pencil, Trash2, Clock, Calendar, Search, MoreHorizontal,
} from '../components/ui/Icons';
import { DateInput } from '../components/form/Inputs';
import { useConfirm } from '../components/ui/useConfirm';
import { useSharePointToken } from '../hooks/useRequests';
import { modeLabel } from '../features/forms/checklistForm';
import { linkState } from '../features/forms/links/linkRules';
import { linkActions, endOfDay, dayOf } from '../features/forms/links/linkChanges';
import {
  filterLinks, countLinks, readFilter, fromTill, linkItems, linkSerials, linkSteps,
  linkTime, groupByDay, dayGroupLabel, avatarTone,
} from '../features/forms/links/linkFilter';
import {
  listLinks, cancelLink, setLinkExpiry, reopenLink, deleteLink, linkUrl,
} from '../features/forms/sharepoint/checklistLinks';
import { formatMYT } from '../utils/malaysiaTime';
import { initialsOf } from '../utils/initials';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

/**
 * Every checklist link IT has shared, as a timeline: where each one stands,
 * and what can be done to it from there. Which actions a link offers is
 * `linkActions`' answer (`features/forms/links/linkChanges.js`); this page
 * only draws them.
 *
 * A signed link opens the employee's signed copy, so "show me what Amir
 * signed" is the same button as "show me what Amir was sent".
 */

const DAY = 86400000;
const SEEN_KEY = 'checklistSeen';

// What each query-string choice means when it is left at its default.
const DEFAULTS = { show: 'all', from: 'any', kind: 'any', q: '', open: '' };

const STATE_LABEL = {
  open: 'Waiting',
  busy: 'Being signed',
  signed: 'Signed',
  expired: 'Expired',
  cancelled: 'Cancelled',
};

// One colour family per state: the knot on the timeline and the pill beside it.
const TONE = {
  open: 'waiting',
  busy: 'waiting',
  signed: 'signed',
  expired: 'closed',
  cancelled: 'closed',
};

const ORBS = [
  { show: 'all', tone: 'all', label: 'Everything', hint: 'Every checklist sent' },
  { show: 'waiting', tone: 'waiting', label: 'Waiting', hint: 'Not signed yet' },
  { show: 'signed', tone: 'signed', label: 'Signed', hint: 'Saved with a signature' },
  { show: 'closed', tone: 'closed', label: 'Closed', hint: 'Expired or cancelled' },
];

const SOURCE_OPTIONS = [
  { value: 'any', label: 'Both' },
  { value: 'till', label: 'From the till' },
  { value: 'share', label: 'Shared by hand' },
];

const KIND_OPTIONS = [
  { value: 'any', label: 'All kinds' },
  { value: 'handout', label: 'Hand outs' },
  { value: 'return', label: 'Returns' },
];

// The actions that can lead the detail panel, in order of preference.
const PRIMARY_ORDER = ['view', 'open', 'reopen'];

const when = (value, style = 'date') => {
  const time = Date.parse(value);
  return Number.isFinite(time) ? formatMYT(time, style) : '—';
};

const keyOf = (link) => String(link.id ?? link.code);

const firstName = (name) => String(name || 'the employee').trim().split(/\s+/)[0];

// The "seen" marks decide which freshly signed rows wear a ring. Kept per
// browser; a private window that refuses storage just shows the ring again.
function readSeen() {
  try {
    const value = JSON.parse(window.localStorage.getItem(SEEN_KEY) ?? '[]');
    return Array.isArray(value) ? value.map(String) : [];
  } catch {
    return [];
  }
}

function writeSeen(list) {
  try {
    window.localStorage.setItem(SEEN_KEY, JSON.stringify(list));
  } catch {
    // Nothing to do: the ring simply comes back next visit.
  }
}

function CopyButton({ url }) {
  const [copied, setCopied] = useState(false);
  const copy = async () => {
    try {
      await navigator.clipboard.writeText(url);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      window.prompt('Copy this link:', url);
    }
  };
  return (
    <button type="button" className="cl-round" aria-label="Copy link" title="Copy link" onClick={copy}>
      {copied ? <Check size={16} /> : <Copy size={16} />}
    </button>
  );
}

/** The date picker for "Change expiry" and "Reopen", opened inside the detail card. */
function DatePanel({ id, label, help, initial, today, confirmLabel, onConfirm, onClose, busy }) {
  const [day, setDay] = useState(initial);
  const valid = Number.isFinite(endOfDay(day)) && day >= today;

  return (
    <div className="cl-panel">
      <label className="ff-label" htmlFor={id}>{label}</label>
      <p className="ff-help">{help}</p>
      <div className="cl-panel-row">
        <DateInput id={id} value={day} onChange={setDay} min={today} />
        <Button size="sm" onClick={() => onConfirm(endOfDay(day))} disabled={!valid || busy}>
          {busy ? 'Saving…' : confirmLabel}
        </Button>
        <Button variant="ghost" size="sm" onClick={onClose} disabled={busy}>Close</Button>
      </div>
    </div>
  );
}

export default function ChecklistLinksPage() {
  const navigate = useNavigate();
  const { instance } = useMsal();
  const getToken = useSharePointToken();
  const { ask, dialog } = useConfirm();
  // Every choice lives in the query string, so a filtered view survives a
  // reload and the till can link straight to its own checklists.
  const [params, setParams] = useSearchParams();
  const filter = readFilter(params);
  const openKey = params.get('open') ?? '';

  // Writes one or more choices; a value at its default is left out of the URL.
  const update = (changes) => {
    const next = new URLSearchParams(params);
    for (const [key, value] of Object.entries(changes)) {
      if (value === DEFAULTS[key]) next.delete(key);
      else next.set(key, value);
    }
    setParams(next, { replace: true });
  };

  const [links, setLinks] = useState(null);
  const [failure, setFailure] = useState('');
  const [now, setNow] = useState(() => Date.now());
  // One row at a time: which link has a date panel open, and which is saving.
  const [panel, setPanel] = useState(null);
  const [working, setWorking] = useState(null);
  const [refreshing, setRefreshing] = useState(false);
  // Which link's "More" menu is open, by its key.
  const [menuFor, setMenuFor] = useState(null);
  const [seen, setSeen] = useState(readSeen);

  useEffect(() => {
    document.title = 'PMW IT — Shared checklists';
  }, []);

  const fetchLinks = useCallback(async () => {
    const tokenRes = await getToken();
    return listLinks({ siteUrl: SHAREPOINT_SITE_URL, token: tokenRes.accessToken });
  }, [getToken]);

  const failed = (thrown) => {
    setFailure(thrown.message || 'The shared checklists could not be loaded');
    setLinks((current) => current ?? []);
  };

  const load = async () => {
    setFailure('');
    try {
      const found = await fetchLinks();
      setLinks(found);
      setNow(() => Date.now());
    } catch (thrown) {
      failed(thrown);
    }
  };

  const refresh = async () => {
    setRefreshing(true);
    try {
      await load();
    } finally {
      setRefreshing(false);
    }
  };

  useEffect(() => {
    let live = true;
    fetchLinks().then(
      (found) => { if (live) setLinks(found); },
      (thrown) => { if (live) failed(thrown); },
    );
    return () => { live = false; };
  }, [fetchLinks]);

  /** Runs one change against one link, then re-reads the list. */
  const run = async (link, change, failureText) => {
    setWorking(link.id);
    setFailure('');
    try {
      const tokenRes = await getToken();
      await change({ siteUrl: SHAREPOINT_SITE_URL, token: tokenRes.accessToken });
      setPanel(null);
      await load();
    } catch (thrown) {
      setFailure(thrown.message || failureText);
    } finally {
      setWorking(null);
    }
  };

  const who = (link) => link.employeeName || 'The employee';

  const cancel = async (link) => {
    const sure = await ask({
      title: 'Cancel this link?',
      body: `${who(link)} will no longer be able to open or sign it. You can reopen it later.`,
      confirmLabel: 'Cancel link',
      cancelLabel: 'Keep it',
    });
    if (sure) run(link, (ctx) => cancelLink({ ...ctx, id: link.id }), 'The link could not be cancelled');
  };

  const expireNow = async (link) => {
    const sure = await ask({
      title: 'Expire this link now?',
      body: `${who(link)} will no longer be able to open it. You can change the date or reopen it later.`,
      confirmLabel: 'Expire now',
      cancelLabel: 'Keep it open',
    });
    if (sure) {
      run(link, (ctx) => setLinkExpiry({ ...ctx, id: link.id, expiresAt: Date.now() }), 'The link could not be expired');
    }
  };

  const remove = async (link) => {
    // Judged by whether a signed record EXISTS, not by the state: a link that
    // was signed and then reopened or expired still has one, and Delete takes
    // it too.
    const signed = Boolean(link.checklistId);
    const sure = await ask({
      title: signed ? 'Delete this signed checklist?' : 'Delete this link?',
      body: signed
        ? `The link, ${who(link)}'s signed checklist and their signature go to the IT helpdesk site's recycle bin. They can be restored from there for 93 days.`
        : 'The link goes to the IT helpdesk site\'s recycle bin and stops opening.',
      confirmLabel: 'Delete',
      cancelLabel: 'Keep it',
    });
    if (sure) run(link, (ctx) => deleteLink({ ...ctx, link }), 'It could not be deleted');
  };

  const openLink = (link) => window.open(linkUrl(window.location.origin, link.code), '_blank', 'noopener');

  /** Picks a row (or, with null, closes the detail) and marks a new link as seen. */
  const choose = (link) => {
    setMenuFor(null);
    if (!link) {
      update({ open: '' });
      return;
    }
    const key = keyOf(link);
    if (!seen.includes(key)) {
      const next = [...seen, key];
      setSeen(next);
      writeSeen(next);
    }
    update({ open: key });
  };

  /** The label, icon and handler of one action. Handlers are the ones this page always had. */
  const specOf = (action, link) => {
    switch (action) {
      case 'open':
        return { label: 'Open', onClick: () => openLink(link) };
      case 'view':
        return { label: 'View signed copy', onClick: () => openLink(link) };
      case 'edit':
        return { label: 'Edit', icon: Pencil, onClick: () => navigate(`/asset-checklist/links/${link.id}`) };
      case 'expiry':
        return { label: 'Change expiry', icon: Calendar, onClick: () => setPanel({ id: link.id, kind: 'expiry' }) };
      case 'expireNow':
        return { label: 'Expire now', icon: Clock, onClick: () => expireNow(link) };
      case 'reopen':
        return { label: 'Reopen', icon: RefreshCw, onClick: () => setPanel({ id: link.id, kind: 'reopen' }) };
      case 'cancel':
        return { label: 'Cancel', icon: X, onClick: () => cancel(link) };
      case 'delete':
        return { label: 'Delete', icon: Trash2, danger: true, onClick: () => remove(link) };
      default:
        return null;
    }
  };

  /** One action as a button: the primary pill, or an item in the More menu. */
  const renderButton = (action, link, role) => {
    const spec = specOf(action, link);
    if (!spec) return null;
    const busy = working === link.id;
    const onClick = role === 'menu'
      ? () => { setMenuFor(null); spec.onClick(); }
      : spec.onClick;

    if (role === 'primary') {
      return (
        <Button key={action} className="cl-primary" disabled={busy} onClick={onClick}>
          {spec.label}
        </Button>
      );
    }
    return (
      <Button
        key={action}
        variant="ghost"
        icon={spec.icon}
        disabled={busy}
        className={spec.danger ? 'cl-danger' : ''}
        onClick={onClick}
      >
        {spec.label}
      </Button>
    );
  };

  const renderPanel = (link) => {
    if (panel?.id !== link.id) return null;
    const busy = working === link.id;

    if (panel.kind === 'expiry') {
      return (
        <DatePanel
          id={`expiry-${link.id}`}
          label="Link works until"
          help="It stops opening at the end of this day."
          initial={dayOf(link.expiresOn) || dayOf(now + 14 * DAY)}
          today={dayOf(now)}
          confirmLabel="Save date"
          busy={busy}
          onClose={() => setPanel(null)}
          onConfirm={(expiresAt) => run(
            link,
            (ctx) => setLinkExpiry({ ...ctx, id: link.id, expiresAt }),
            'The date could not be changed',
          )}
        />
      );
    }

    const wasSigned = Boolean(link.checklistId);
    return (
      <DatePanel
        id={`reopen-${link.id}`}
        label="Reopen until"
        help={wasSigned
          ? `${who(link)} gets the same link back, filled with what they signed. They can keep their signature or sign again, and the same record is updated.`
          : `${who(link)} can open and sign the same link again.`}
        initial={dayOf(now + 14 * DAY)}
        today={dayOf(now)}
        confirmLabel="Reopen"
        busy={busy}
        onClose={() => setPanel(null)}
        onConfirm={(expiresAt) => run(
          link,
          (ctx) => reopenLink({ ...ctx, link, expiresAt }),
          'The link could not be reopened',
        )}
      />
    );
  };

  // The signed-in IT person, for nothing but the "edited by" wording below.
  const me = instance.getActiveAccount()?.name;

  /** One row on the timeline. */
  const renderRow = (link) => {
    const key = keyOf(link);
    const state = linkState(link, now);
    const name = link.employeeName || 'No name set';
    const time = linkTime(link);
    const isNew = state === 'signed' && Number.isFinite(time) && now - time < DAY && !seen.includes(key);
    const items = linkItems(link);
    const tillLink = fromTill(link);
    const today = dayGroupLabel(time, now) === 'Today';
    const selected = openKey === key;

    return (
      <button
        key={key}
        type="button"
        className={`cl-trow${selected ? ' on' : ''}`}
        aria-pressed={selected}
        onClick={() => choose(link)}
      >
        <span className={`cl-knot cl-knot-${TONE[state]}`} aria-hidden="true" />
        <span className={`cl-av cl-av-${avatarTone(name)}${isNew ? ' cl-av-new' : ''}`} aria-hidden="true">
          {initialsOf(link.employeeName)}
        </span>
        <span className="cl-trow-main">
          <span className="cl-trow-name">{name}</span>
          <span className="cl-trow-meta">
            {modeLabel(link.formMode)}{tillLink ? ' · From the till' : ''}
          </span>
          {items.length > 0 && (
            <span className="cl-chips">
              {items.slice(0, 3).map((item) => (
                <span key={item.name} className="cl-minichip"><b>{item.qty}</b>{item.name}</span>
              ))}
              {items.length > 3 && <span className="cl-minichip">+{items.length - 3} more</span>}
            </span>
          )}
        </span>
        <span className="cl-trow-end">
          <span className={`cl-pill cl-pill-${TONE[state]}`}>
            <i className="cl-dot" aria-hidden="true" />
            {STATE_LABEL[state]}
          </span>
          <span className="cl-time">{Number.isFinite(time) ? formatMYT(time, today ? 'time' : 'date') : '—'}</span>
        </span>
      </button>
    );
  };

  /** The detail panel for the selected link: its story, what it covers, and what can be done. */
  const renderDetail = (link) => {
    const state = linkState(link, now);
    const name = link.employeeName || 'No name set';
    const tillLink = fromTill(link);
    const items = linkItems(link);
    const serials = linkSerials(link);
    const steps = linkSteps(link, now);
    const acts = linkActions(link, now);
    const primary = PRIMARY_ORDER.find((action) => acts.includes(action));
    const menu = acts.filter((action) => action !== 'copy' && action !== primary);
    const menuOpen = menuFor === keyOf(link);
    const handoverCount = link.handovers?.ids?.length ?? 0;
    const handoverWord = link.handovers?.kind === 'return' ? 'return' : 'handover';

    return (
      <>
        <section className="cl-dhead">
          <span className="cl-deco cl-deco-a" aria-hidden="true" />
          <span className="cl-deco cl-deco-b" aria-hidden="true" />
          <div className="cl-dhead-top">
            <span className="cl-dhead-av" aria-hidden="true">{initialsOf(link.employeeName)}</span>
            <div className="cl-dhead-text">
              <h2 className="cl-dhead-name">{name}</h2>
              <p className="cl-dhead-sub">
                {modeLabel(link.formMode)} · {tillLink ? 'from the till' : 'shared by hand'}
              </p>
            </div>
            <button type="button" className="cl-round cl-round-light" aria-label="Close" onClick={() => choose(null)}>
              <X size={16} />
            </button>
          </div>
          <div className="cl-steps">
            {steps.map((step, index) => (
              <Fragment key={`${step.label}-${index}`}>
                {index > 0 && (
                  <span className={`cl-step-line${step.done ? ' cl-step-line-done' : ''}`} aria-hidden="true" />
                )}
                <div className={`cl-step${step.done ? ' cl-step-done' : ''}`}>
                  <span className="cl-step-dot" aria-hidden="true">{step.done && <Check size={16} />}</span>
                  <b>{step.label}</b>
                  <small>{step.when == null ? 'Not yet' : formatMYT(step.when, 'datetime12')}</small>
                </div>
              </Fragment>
            ))}
          </div>
        </section>

        <div className="cl-dcard">
          {items.length > 0 && (
            <div className="cl-dsec">
              <h3 className="cl-dsec-h">What it covers</h3>
              <div className="cl-chips">
                {items.map((item) => (
                  <span key={item.name} className="cl-minichip"><b>{item.qty}</b>{item.name}</span>
                ))}
              </div>
            </div>
          )}

          {serials.length > 0 && (
            <div className="cl-dsec">
              <h3 className="cl-dsec-h">Serial numbers</h3>
              <ul className="cl-serials">
                {serials.map((line, index) => (
                  <li key={index}>
                    <i className="cl-dot" aria-hidden="true" />
                    {line.what && <span>{line.what}</span>}
                    <code className="cl-serial">{line.serial}</code>
                  </li>
                ))}
              </ul>
            </div>
          )}

          {state === 'signed' && (
            <p className="cl-strip cl-strip-good">
              Signed {when(link.signedOn, 'datetime12')}
              {tillLink && ` · added to ${handoverCount} ${handoverWord} records`}
            </p>
          )}

          {(state === 'open' || state === 'busy') && (
            <p className="cl-strip cl-strip-warn">
              Waiting for {firstName(link.employeeName)} · link works until {when(link.expiresOn)}
            </p>
          )}

          {link.editedBy && state === 'signed' && (
            <p className="cl-edited">
              Edited after signing by {link.editedBy === me ? 'you' : link.editedBy}, {when(link.editedOn, 'datetime12')}
            </p>
          )}

          {renderPanel(link)}

          <div className="cl-dactions">
            {acts.includes('copy') && <CopyButton url={linkUrl(window.location.origin, link.code)} />}
            {primary && renderButton(primary, link, 'primary')}
            {menu.length > 0 && (
              <div className="cl-more">
                <button
                  type="button"
                  className="cl-round"
                  aria-label="More actions"
                  aria-expanded={menuOpen}
                  onClick={() => setMenuFor(menuOpen ? null : keyOf(link))}
                >
                  <MoreHorizontal size={18} />
                </button>
                {menuOpen && (
                  <div className="cl-menu">
                    {menu.map((action) => renderButton(action, link, 'menu'))}
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      </>
    );
  };

  const actions = (
    <>
      <label className="cl-find">
        <Search size={16} aria-hidden="true" />
        <input
          type="search"
          placeholder="Search name, item or serial"
          aria-label="Search checklists"
          value={filter.q}
          onChange={(event) => update({ q: event.target.value })}
        />
      </label>
      <button
        type="button"
        className="cl-round"
        aria-label="Refresh"
        title="Refresh"
        disabled={refreshing}
        onClick={refresh}
      >
        {refreshing ? <Spinner size={14} /> : <RefreshCw size={16} />}
      </button>
      <button type="button" className="cl-new" onClick={() => navigate('/asset-checklist/share')}>
        <span className="cl-new-dot"><Plus size={14} /></span>
        New checklist
      </button>
    </>
  );

  const shown = filterLinks(links, filter, now);
  const counts = countLinks(links, filter, now);
  const groups = groupByDay(shown, now);
  const selected = (links ?? []).find((link) => keyOf(link) === openKey) ?? null;

  return (
    <AppShell
      title="Checklists"
      subtitle="What people signed for, and what is still waiting"
      actions={actions}
    >
      {failure && <ErrorBanner message={failure} onRetry={load} />}

      {links === null && (
        <Card className="ff-progress"><span className="spinner" /> Loading…</Card>
      )}

      {links?.length === 0 && !failure && (
        <EmptyState>
          No checklist has been shared yet. <strong>New checklist</strong> to make the first link.
        </EmptyState>
      )}

      {links?.length > 0 && (
        <>
          <div className="cl-orbs">
            {ORBS.map((orb) => (
              <button
                key={orb.show}
                type="button"
                className="cl-orb"
                aria-pressed={filter.show === orb.show}
                onClick={() => update({ show: orb.show })}
              >
                <span className={`cl-orb-count cl-orb-${orb.tone}`}>{counts[orb.show]}</span>
                <span className="cl-orb-text">
                  <strong>{orb.label}</strong>
                  <small>{orb.hint}</small>
                </span>
              </button>
            ))}
          </div>

          <div className="cl-split">
            <section className="cl-board" aria-label="Checklists">
              <div className="cl-seg-row">
                <div className="cl-seg" role="group" aria-label="Where they were made">
                  {SOURCE_OPTIONS.map((option) => (
                    <button
                      key={option.value}
                      type="button"
                      aria-pressed={filter.source === option.value}
                      onClick={() => update({ from: option.value })}
                    >
                      {option.label}
                    </button>
                  ))}
                </div>
                <div className="cl-seg" role="group" aria-label="What kind">
                  {KIND_OPTIONS.map((option) => (
                    <button
                      key={option.value}
                      type="button"
                      aria-pressed={filter.kind === option.value}
                      onClick={() => update({ kind: option.value })}
                    >
                      {option.label}
                    </button>
                  ))}
                </div>
              </div>

              {shown.length === 0 ? (
                <div className="cl-clear">
                  <span className="cl-clear-orb"><Check size={36} /></span>
                  <strong>All clear</strong>
                  <span>Nothing matches these filters.</span>
                </div>
              ) : groups.map((group) => (
                <div className="cl-day" key={group.label}>
                  <div className="cl-day-label">{group.label}</div>
                  <div className="cl-track">
                    {group.links.map(renderRow)}
                  </div>
                </div>
              ))}
            </section>

            {selected ? (
              <>
                <button type="button" className="cl-scrim" aria-label="Close details" onClick={() => choose(null)} />
                <aside className="cl-detail" aria-label="Checklist details">
                  {renderDetail(selected)}
                </aside>
              </>
            ) : (
              <aside className="cl-detail cl-detail-idle">
                Choose a checklist to see its steps and what can be done with it.
              </aside>
            )}
          </div>
        </>
      )}

      {dialog}
    </AppShell>
  );
}

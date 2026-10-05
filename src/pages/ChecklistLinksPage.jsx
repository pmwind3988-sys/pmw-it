import { useCallback, useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card, ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import {
  Copy, Check, Plus, X, RefreshCw, Pencil, Trash2, Clock, Calendar,
} from '../components/ui/Icons';
import { DateInput } from '../components/form/Inputs';
import { useConfirm } from '../components/ui/useConfirm';
import { useSharePointToken } from '../hooks/useRequests';
import { modeLabel } from '../features/forms/checklistForm';
import { linkState } from '../features/forms/links/linkRules';
import { linkActions, endOfDay, dayOf } from '../features/forms/links/linkChanges';
import {
  listLinks, cancelLink, setLinkExpiry, reopenLink, deleteLink, linkUrl,
} from '../features/forms/sharepoint/checklistLinks';
import { formatMYT } from '../utils/malaysiaTime';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

/**
 * Every checklist link IT has shared, where each one stands, and what can be
 * done to it from there. Which actions a link offers is `linkActions`' answer
 * (`features/forms/links/linkChanges.js`); this page only draws them.
 *
 * A signed link opens the employee's signed copy, so "show me what Amir
 * signed" is the same button as "show me what Amir was sent".
 */

const STATE_LABEL = {
  open: 'Waiting',
  busy: 'Being signed',
  signed: 'Signed',
  expired: 'Expired',
  cancelled: 'Cancelled',
};

const DAY = 86400000;

const when = (value, style = 'date') => {
  const time = Date.parse(value);
  return Number.isFinite(time) ? formatMYT(time, style) : '—';
};

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
    <Button variant="ghost" size="sm" icon={copied ? Check : Copy} onClick={copy}>
      {copied ? 'Copied' : 'Copy link'}
    </Button>
  );
}

/** The date picker for "Change expiry" and "Reopen", opened inside the row. */
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

  const [links, setLinks] = useState(null);
  const [failure, setFailure] = useState('');
  const [now, setNow] = useState(() => Date.now());
  // One row at a time: which link has a date panel open, and which is saving.
  const [panel, setPanel] = useState(null);
  const [working, setWorking] = useState(null);

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

  const open = (link) => window.open(linkUrl(window.location.origin, link.code), '_blank', 'noopener');

  const renderAction = (action, link) => {
    const busy = working === link.id;
    const common = { variant: 'ghost', size: 'sm', disabled: busy };
    switch (action) {
      case 'copy':
        return <CopyButton key={action} url={linkUrl(window.location.origin, link.code)} />;
      case 'open':
        return <Button key={action} {...common} onClick={() => open(link)}>Open</Button>;
      case 'view':
        return <Button key={action} {...common} onClick={() => open(link)}>View signed copy</Button>;
      case 'edit':
        return (
          <Button key={action} {...common} icon={Pencil} onClick={() => navigate(`/asset-checklist/links/${link.id}`)}>
            Edit
          </Button>
        );
      case 'expiry':
        return (
          <Button key={action} {...common} icon={Calendar} onClick={() => setPanel({ id: link.id, kind: 'expiry' })}>
            Change expiry
          </Button>
        );
      case 'expireNow':
        return <Button key={action} {...common} icon={Clock} onClick={() => expireNow(link)}>Expire now</Button>;
      case 'reopen':
        return (
          <Button key={action} {...common} icon={RefreshCw} onClick={() => setPanel({ id: link.id, kind: 'reopen' })}>
            Reopen
          </Button>
        );
      case 'cancel':
        return <Button key={action} {...common} icon={X} onClick={() => cancel(link)}>Cancel</Button>;
      case 'delete':
        return (
          <Button key={action} {...common} icon={Trash2} className="cl-danger" onClick={() => remove(link)}>
            Delete
          </Button>
        );
      default:
        return null;
    }
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

  const actions = (
    <>
      <Button variant="ghost" icon={RefreshCw} onClick={load}>Refresh</Button>
      <Button icon={Plus} onClick={() => navigate('/asset-checklist/share')}>Share a checklist</Button>
    </>
  );

  // The signed-in IT person, for nothing but the "edited by" wording below.
  const me = instance.getActiveAccount()?.name;

  return (
    <AppShell
      title="Shared checklists"
      subtitle="Links sent to employees to fill in and sign, newest first"
      actions={actions}
    >
      {failure && <ErrorBanner message={failure} onRetry={load} />}

      {links === null && (
        <Card className="ff-progress"><span className="spinner" /> Loading…</Card>
      )}

      {links?.length === 0 && !failure && (
        <EmptyState>
          No checklist has been shared yet. <strong>Share a checklist</strong> to make the first link.
        </EmptyState>
      )}

      {links?.length > 0 && (
        <ul className="cl-list">
          {links.map((link) => {
            const state = linkState(link, now);
            return (
              <li key={link.id ?? link.code}>
                <Card className="cl-row">
                  <div className="cl-row-main">
                    <span className="cl-row-name">{link.employeeName || 'No name set'}</span>
                    <span className="cl-row-meta">
                      {modeLabel(link.formMode)} · shared {when(link.created)}
                      {link.createdByName ? ` by ${link.createdByName}` : ''}
                    </span>
                    <span className="cl-row-meta">
                      {state === 'signed'
                        ? `Signed ${when(link.signedOn, 'datetime12')}`
                        : `${state === 'expired' ? 'Expired' : 'Expires'} ${when(link.expiresOn)}`}
                    </span>
                    {link.editedBy && state === 'signed' && (
                      <span className="cl-row-meta cl-row-edited">
                        Edited after signing by {link.editedBy === me ? 'you' : link.editedBy}, {when(link.editedOn, 'datetime12')}
                      </span>
                    )}
                  </div>
                  <span className={`cl-state cl-state-${state}`}>{STATE_LABEL[state]}</span>
                  <div className="cl-row-actions">
                    {linkActions(link, now).map((action) => renderAction(action, link))}
                  </div>
                  {renderPanel(link)}
                </Card>
              </li>
            );
          })}
        </ul>
      )}

      {dialog}
    </AppShell>
  );
}

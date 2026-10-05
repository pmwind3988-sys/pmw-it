import { useCallback, useEffect, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card, ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import { Copy, Check, Plus, X, RefreshCw } from '../components/ui/Icons';
import { useConfirm } from '../components/ui/useConfirm';
import { useSharePointToken } from '../hooks/useRequests';
import { modeLabel } from '../features/forms/checklistForm';
import { linkState } from '../features/forms/links/linkRules';
import { listLinks, cancelLink, linkUrl } from '../features/forms/sharepoint/checklistLinks';
import { formatMYT } from '../utils/malaysiaTime';

const SHAREPOINT_SITE_URL =
  import.meta.env.VITE_SHAREPOINT_SITE_URL || 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

/**
 * Every checklist link IT has shared, and where each one stands.
 *
 * A signed link opens the employee's signed copy, so "show me what Amir
 * signed" is the same Open button as "show me what Amir was sent".
 */

const STATE_LABEL = {
  open: 'Waiting',
  busy: 'Being signed',
  signed: 'Signed',
  expired: 'Expired',
  cancelled: 'Cancelled',
};

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

export default function ChecklistLinksPage() {
  const navigate = useNavigate();
  const getToken = useSharePointToken();
  const { ask, dialog } = useConfirm();

  const [links, setLinks] = useState(null);
  const [failure, setFailure] = useState('');
  const [now] = useState(() => Date.now());

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
      setLinks(await fetchLinks());
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

  const cancel = async (link) => {
    const sure = await ask({
      title: 'Cancel this link?',
      body: `${link.employeeName || 'The employee'} will no longer be able to open or sign it. This cannot be undone — you would share a new one instead.`,
      confirmLabel: 'Cancel link',
      cancelLabel: 'Keep it',
    });
    if (!sure) return;

    try {
      const tokenRes = await getToken();
      await cancelLink({ siteUrl: SHAREPOINT_SITE_URL, token: tokenRes.accessToken, id: link.id });
      await load();
    } catch (thrown) {
      setFailure(thrown.message || 'The link could not be cancelled');
    }
  };

  const actions = (
    <>
      <Button variant="ghost" icon={RefreshCw} onClick={load}>Refresh</Button>
      <Button icon={Plus} onClick={() => navigate('/asset-checklist/share')}>Share a checklist</Button>
    </>
  );

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
            const url = linkUrl(window.location.origin, link.code);
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
                        : `Expires ${when(link.expiresOn)}`}
                    </span>
                  </div>
                  <span className={`cl-state cl-state-${state}`}>{STATE_LABEL[state]}</span>
                  <div className="cl-row-actions">
                    {(state === 'open' || state === 'busy') && <CopyButton url={url} />}
                    {state !== 'cancelled' && state !== 'expired' && (
                      <Button variant="ghost" size="sm" onClick={() => window.open(url, '_blank', 'noopener')}>
                        {state === 'signed' ? 'View signed copy' : 'Open'}
                      </Button>
                    )}
                    {state === 'open' && (
                      <Button variant="ghost" size="sm" icon={X} onClick={() => cancel(link)}>
                        Cancel
                      </Button>
                    )}
                  </div>
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

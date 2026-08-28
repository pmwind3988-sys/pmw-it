import { useMemo, useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card, ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import {
  ArrowLeft, Check, AlertTriangle, Clock, Package, Users,
} from '../components/ui/Icons';
import { useHandovers } from '../features/assets/useHandovers';
import { SHAREPOINT_SITE_URL } from '../features/assets/useAssets';
import { useSharePointToken } from '../hooks/useRequests';
import { commitReturn } from '../features/assets/sharepoint/writeHandover';
import SignatureField from '../features/assets/ui/SignatureField';
import { absoluteFileUrl } from '../features/assets/sharepoint/fileUrl';
import { returnEverything } from '../features/assets/handover/planReturn';
import {
  heldBy, outstanding, isOverdue, isOpen, HANDOVER_KIND,
} from '../features/assets/handover/availability';
import { CONDITIONS } from '../features/assets/assetKinds';
import { formatMYT } from '../utils/malaysiaTime';
import { initialsOf } from '../utils/initials';

/**
 * One person and everything they hold.
 *
 * Keyed on email rather than on a name in the URL, because the email is the
 * identity everything per-person hangs off (§4.5) — and a name in a path would
 * break the moment somebody's changed.
 */
/**
 * A signature that was captured, as a link to the picture of it. Nothing at
 * all when there is none -- a handover recorded without one is a normal thing
 * and does not need a row saying "unsigned" against it.
 */
function Signed({ stored, what }) {
  if (!stored) return null;

  return (
    <a
      href={absoluteFileUrl(SHAREPOINT_SITE_URL, stored)}
      target="_blank"
      rel="noreferrer"
      className="as-signed"
      title={`Signed for ${what}`}
    >
      <Check size={11} /> signed
    </a>
  );
}

export default function AssetPersonPage() {
  const { email } = useParams();
  const navigate = useNavigate();
  const { instance } = useMsal();
  const getToken = useSharePointToken();
  const { handovers, loading, error, reload } = useHandovers();

  const [condition, setCondition] = useState(CONDITIONS[1] ?? 'Good');
  // Signed for on the way back in, the same as on the way out. Cleared after
  // each return: one signature covers what was just handed back, not the next
  // thing somebody carries in an hour later.
  const [signature, setSignature] = useState(null);
  const [busy, setBusy] = useState(false);
  const [report, setReport] = useState(null);
  const [failure, setFailure] = useState('');

  const address = decodeURIComponent(email ?? '');
  const held = useMemo(() => heldBy(handovers, address), [handovers, address]);
  const past = useMemo(
    () => handovers
      .filter((row) => String(row.personEmail ?? '').toLowerCase() === address.toLowerCase())
      .filter((row) => !isOpen(row))
      .sort((a, b) => (b.returnedOn ?? 0) - (a.returnedOn ?? 0)),
    [handovers, address],
  );

  const name = held[0]?.personName || past[0]?.personName || address;
  const overdue = held.filter((row) => isOverdue(row)).length;
  const units = held.reduce((sum, row) => sum + outstanding(row), 0);

  const doReturn = async (entries) => {
    setBusy(true);
    setFailure('');
    setReport(null);
    try {
      const tokenRes = await getToken();
      const account = instance.getActiveAccount();
      const result = await commitReturn({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        returns: entries,
        returnedBy: account?.username ?? account?.name ?? '',
        signature,
      });
      setReport(result);
      setSignature(null);
      reload();
    } catch (thrown) {
      setFailure(thrown.message || 'The return could not be recorded');
    } finally {
      setBusy(false);
    }
  };

  if (loading) return <AppShell title="Person"><div className="spinner" /></AppShell>;

  return (
    <AppShell
      title={name}
      subtitle={address}
      actions={(
        <>
          <Button variant="ghost" icon={ArrowLeft} onClick={() => navigate('/assets/people')}>
            Everyone
          </Button>
          {held.length > 0 && (
            <Button
              icon={Check}
              disabled={busy}
              onClick={() => doReturn(returnEverything(held, condition))}
            >
              {busy ? 'Returning…' : 'Return everything'}
            </Button>
          )}
        </>
      )}
    >
      {error && <ErrorBanner message={error} onRetry={reload} />}
      {failure && <ErrorBanner message={failure} />}

      {report && (
        <Card className={`as-notice ${report.blocked.length ? 'as-notice-warn' : 'as-notice-ok'}`}>
          {report.blocked.length ? <AlertTriangle size={16} /> : <Check size={16} />}
          <span>
            {report.returned} line{report.returned === 1 ? '' : 's'} returned.
            {report.blocked.length > 0
              && ` ${report.blocked.length} could not be: ${report.blocked[0].reason}`}
            {report.signed && ' Signed for.'}
            {report.signatureFailed
              && ' The signature could not be saved, so it is recorded unsigned.'}
          </span>
        </Card>
      )}

      <div className="as-personhead">
        <span className="as-avatar as-avatar-lg">{initialsOf(name)}</span>
        <div className="as-facts as-facts-inline">
          <span><strong>{units}</strong> item{units === 1 ? '' : 's'} held</span>
          {overdue > 0 && (
            <span className="as-overdue">
              <Clock size={13} /> <strong>{overdue}</strong> overdue
            </span>
          )}
        </div>
      </div>

      {held.length > 0 && (
        <Card className="as-panel">
          <h2 className="as-h2">Coming back in what condition?</h2>
          <p className="as-hint">
            Applied to whatever you return next. A faulty item must not rejoin the
            shelf looking available.
          </p>
          <select
            className="as-conditionpick"
            value={condition}
            onChange={(event) => setCondition(event.target.value)}
            aria-label="Condition on return"
          >
            {CONDITIONS.map((entry) => <option key={entry} value={entry}>{entry}</option>)}
          </select>

          {/* Signed on the way in as well as on the way out. "It came back" is
              a claim about somebody else's property, and the person handing it
              back is the one who can stand behind it. */}
          <SignatureField
            label={`${name} handed these back`}
            hint={'Recommended — ask them to sign for what they are returning. '
              + 'It can be skipped.'}
            value={signature}
            onChange={setSignature}
            disabled={busy}
          />
        </Card>
      )}

      <h2 className="as-h2">Currently holding</h2>
      {held.length === 0 ? (
        <EmptyState>
          <Package size={20} />
          {name} has nothing out.
        </EmptyState>
      ) : (
        <div className="as-table-wrap">
          <table className="as-table">
            <thead>
              <tr>
                <th>Item</th>
                <th>Out</th>
                <th>Kind</th>
                <th>Since</th>
                <th>Due</th>
                <th />
              </tr>
            </thead>
            <tbody>
              {held.map((row) => (
                <tr key={row.id} className={isOverdue(row) ? 'as-row-overdue' : undefined}>
                  <td>
                    <Link to={`/assets/${row.assetId}`} className="as-link">{row.itemTitle}</Link>
                    <span className="as-sub">
                      {row.category}
                      {row.serialNumber && ` · ${row.serialNumber}`}
                    </span>
                  </td>
                  <td className="as-qty">{outstanding(row)}</td>
                  <td>{row.kind}</td>
                  <td className="as-when">
                    {row.issuedOnMYT || '—'}
                    <Signed stored={row.issueSignature} what="handing this over" />
                  </td>
                  <td className="as-when">
                    {row.kind === HANDOVER_KIND.BORROWED
                      ? (row.dueOnMYT || (row.dueOn ? formatMYT(row.dueOn, 'datetime12') : '—'))
                      : '—'}
                  </td>
                  <td>
                    <Button
                      variant="ghost"
                      size="sm"
                      disabled={busy}
                      onClick={() => doReturn([{
                        handoverId: row.id, quantity: outstanding(row), condition,
                      }])}
                    >
                      Return
                    </Button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {past.length > 0 && (
        <details className="as-batchlist">
          <summary><Users size={13} /> Already returned ({past.length})</summary>
          <ul>
            {past.map((row) => (
              <li key={row.id}>
                <Link to={`/assets/${row.assetId}`} className="as-link">{row.itemTitle}</Link>
                <span className="as-sub">
                  {row.quantity} back {row.returnedOnMYT || ''}
                  {row.returnCondition ? ` — ${row.returnCondition}` : ''}
                  <Signed stored={row.issueSignature} what="taking it" />
                  <Signed stored={row.returnSignature} what="bringing it back" />
                </span>
              </li>
            ))}
          </ul>
        </details>
      )}
    </AppShell>
  );
}

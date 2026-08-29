import { useMemo, useState } from 'react';
import { Link, useNavigate, useParams, useSearchParams } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card, ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import {
  ArrowLeft, Check, AlertTriangle, Clock, Package, Users, Pencil,
} from '../components/ui/Icons';
import { useHandovers } from '../features/assets/useHandovers';
import { SHAREPOINT_SITE_URL } from '../features/assets/useAssets';
import { useSharePointToken } from '../hooks/useRequests';
import { commitReturn, commitPersonEdit } from '../features/assets/sharepoint/writeHandover';
import SignatureField from '../features/assets/ui/SignatureField';
import SignatureShot from '../features/assets/ui/SignatureShot';
import PersonEditor from '../features/assets/ui/PersonEditor';
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

  // Correcting who this person is. Kept apart from the return above: they are
  // two different actions with two different risks, and one saying "4 lines
  // returned" while the other says "renamed" in the same box would be a
  // sentence nobody could act on.
  // Opened straight into the form when the people list sent us here to fix a
  // name, and closed by ordinary use afterwards — the query says how the page
  // was arrived at, not what state it must stay in.
  const [params] = useSearchParams();
  const [editing, setEditing] = useState(params.get('edit') === '1');
  const [renamed, setRenamed] = useState(null);

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
      // Kept when it did NOT save. Clearing it regardless threw the drawing
      // away at the one moment it still existed, so "the signature could not
      // be saved" was a dead end — now the button can simply be pressed again
      // and the same signature goes up with it.
      if (!result.signatureFailed) setSignature(null);
      reload();
    } catch (thrown) {
      setFailure(thrown.message || 'The return could not be recorded');
    } finally {
      setBusy(false);
    }
  };

  const doRename = async (draft) => {
    setBusy(true);
    setFailure('');
    setRenamed(null);
    setReport(null);
    try {
      const tokenRes = await getToken();
      const person = {
        name: draft.name.trim(),
        email: draft.email.trim().toLowerCase(),
        login: '',
      };

      const result = await commitPersonEdit({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        from: address,
        person,
      });

      setRenamed({ ...result, ...person });
      setEditing(false);
      reload();

      // The page is keyed on the email, so a changed one leaves this URL
      // pointing at somebody who no longer exists. Replaced rather than
      // pushed: going Back should reach the people list, not a dead address.
      if (person.email !== address.toLowerCase()) {
        navigate(`/assets/people/${encodeURIComponent(person.email)}`, { replace: true });
      }
    } catch (thrown) {
      setFailure(thrown.message || 'The details could not be changed');
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
          {!editing && (
            <Button
              variant="secondary"
              icon={Pencil}
              disabled={busy}
              onClick={() => { setEditing(true); setRenamed(null); }}
            >
              Edit details
            </Button>
          )}
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
              && ' The signature could not be saved, so this is recorded unsigned.'
                + ' It is still on the screen below and will go up with the next return.'}
          </span>
        </Card>
      )}

      {renamed && (
        <Card className={`as-notice ${renamed.staleRows.length ? 'as-notice-warn' : 'as-notice-ok'}`}>
          {renamed.staleRows.length ? <AlertTriangle size={16} /> : <Check size={16} />}
          <span>
            Now {renamed.name} · {renamed.email}, across {renamed.changed} handover
            {renamed.changed === 1 ? '' : 's'}.
            {renamed.openLines > 0
              && ` The ${renamed.openLines} item${renamed.openLines === 1 ? '' : 's'} still out `
                + 'stayed exactly where they were.'}
            {renamed.writeFailures.length > 0
              && ` ${renamed.writeFailures.length} row${renamed.writeFailures.length === 1 ? '' : 's'} `
                + 'could not be changed and still read the old details. Try again.'}
            {renamed.staleRows.length > 0
              && ' Some register rows still show the old name; the handovers themselves are correct.'}
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

      {editing && (
        <Card className="as-panel">
          <h2 className="as-h2">Who this is</h2>
          <PersonEditor
            current={{ name, email: address }}
            handovers={handovers}
            busy={busy}
            onSave={doRename}
            onCancel={() => setEditing(false)}
          />
        </Card>
      )}

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
                    {/* This handover's own signature, not a tick shared by
                        every line on the page. Two things taken on two
                        different days were signed for twice, and the register
                        should show both. */}
                    <SignatureShot
                      stored={row.issueSignature}
                      when="signed on the way out"
                      by={row.personName || row.personEmail}
                    />
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
                </span>
                {/* Both halves of the round trip, side by side. Whether the
                    hand that took it is the hand that brought it back is the
                    question this record exists to answer, and it can only be
                    answered by seeing the two. */}
                <span className="as-sigpair">
                  <SignatureShot
                    stored={row.issueSignature}
                    when="signed on the way out"
                    by={row.personName || row.personEmail}
                  />
                  <SignatureShot
                    stored={row.returnSignature}
                    when="signed on the way back"
                    by={row.personName || row.personEmail}
                  />
                </span>
              </li>
            ))}
          </ul>
        </details>
      )}
    </AppShell>
  );
}

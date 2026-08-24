import { useCallback, useMemo, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import { useMsal } from '@azure/msal-react';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card, ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import {
  ScanLine, Search, Plus, Trash2, Check, AlertTriangle, X, Package,
} from '../components/ui/Icons';
import { useAssets, SHAREPOINT_SITE_URL } from '../features/assets/useAssets';
import { useSharePointToken } from '../hooks/useRequests';
import { useScanner, CAMERA_STATE } from '../features/assets/scan/useScanner';
import { signalAccepted, signalDuplicate } from '../features/assets/scan/feedback';
import {
  newBasket, newLine, addLine, removeLine, setQuantity, replaceLine, hasAsset, unitCount,
} from '../features/assets/handover/basket';
import { lineRefusal } from '../features/assets/handover/planHandover';
import { available, owned, HANDOVER_KIND } from '../features/assets/handover/availability';
import { commitHandover } from '../features/assets/sharepoint/writeHandover';
import { matchesQuery } from '../features/assets/assetFilters';
import { normaliseCode } from '../features/assets/identity';
import { TRACKED } from '../features/assets/assetKinds';
import { parseUnits } from '../features/assets/units';
import PersonPicker from '../features/assets/ui/PersonPicker';

/**
 * Handing several things to one person as a single event.
 *
 * The person is picked first on purpose: a line cannot be checked until it is
 * known who it is for, and a refusal that only arrives at the end is one nobody
 * can act on while still standing at the desk (§4.4).
 */

const PHASE_LABEL = {
  provisioning: 'Setting up the SharePoint lists',
  reading: 'Checking what is available',
  writing: 'Recording the handover',
  updating: 'Updating the register',
};

/** `datetime-local` wants the local wall clock, not an ISO instant. */
function toLocalDate(epochMs) {
  if (!Number.isFinite(epochMs)) return '';
  const offset = new Date(epochMs).getTimezoneOffset() * 60000;
  return new Date(epochMs - offset).toISOString().slice(0, 10);
}

export default function AssetHandoverPage() {
  const navigate = useNavigate();
  const { instance } = useMsal();
  const getToken = useSharePointToken();
  const { assets, loading, error, reload } = useAssets();

  const [basket, setBasket] = useState(() => newBasket());
  const [query, setQuery] = useState('');
  const [scanning, setScanning] = useState(false);
  const [flash, setFlash] = useState(null);
  const [progress, setProgress] = useState(null);
  const [report, setReport] = useState(null);
  const [failure, setFailure] = useState('');

  const byId = useMemo(() => new Map(assets.map((asset) => [asset.id, asset])), [assets]);

  /** Only what there is any of. Offering an empty box is offering a refusal. */
  const matches = useMemo(() => {
    if (!query.trim()) return [];
    return assets
      .filter((asset) => available(asset) > 0 && matchesQuery(asset, query))
      .slice(0, 8);
  }, [assets, query]);

  const add = useCallback((asset) => {
    setBasket((current) => {
      if (hasAsset(current, asset.id)) return current;
      return addLine(current, newLine(asset));
    });
    setQuery('');
  }, []);

  /**
   * A scanned code is matched against every code on a row, because whoever is
   * holding the thing does not know which field their barcode ended up in.
   */
  const onCodes = useCallback((codes) => {
    for (const code of codes) {
      const wanted = normaliseCode(code.rawValue);
      const found = assets.find((asset) => (
        normaliseCode(asset.serialNumber) === wanted
        || normaliseCode(asset.assetTag) === wanted
        || normaliseCode(asset.partNumber) === wanted
        || (asset.additionalCodes ?? []).some((entry) => normaliseCode(entry) === wanted)
        // A bulk line keeps its codes on its individual items, so scanning the
        // tab in your hand has to reach them or the register answers "nothing
        // matches" for something it certainly owns.
        || parseUnits(asset.units).some((unit) => (
          normaliseCode(unit.serialNumber) === wanted
          || normaliseCode(unit.assetTag) === wanted
          || normaliseCode(unit.partNumber) === wanted
        ))
      ));

      if (!found) {
        signalDuplicate();
        setFlash({ kind: 'dup', text: `Nothing in the register matches ${code.rawValue}` });
        continue;
      }

      if (available(found) < 1) {
        signalDuplicate();
        setFlash({ kind: 'dup', text: `${found.title} is already out` });
        continue;
      }

      signalAccepted();
      setFlash({ kind: 'ok', text: found.title });
      add(found);
    }
  }, [assets, add]);

  const { videoRef, state } = useScanner({ active: scanning, onCodes });

  const refusals = useMemo(() => new Map(
    basket.lines.map((line) => [line.lineId, lineRefusal(line, byId.get(line.assetId), basket)]),
  ), [basket, byId]);

  const blockedCount = [...refusals.values()].filter(Boolean).length;

  const handOver = async () => {
    setFailure('');
    setReport(null);
    try {
      const tokenRes = await getToken();
      const account = instance.getActiveAccount();

      const result = await commitHandover({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        basket,
        issuedBy: account?.username ?? account?.name ?? '',
        onProgress: setProgress,
      });

      setProgress(null);
      setReport(result);
      reload();

      // Only what was refused stays in the basket, so pressing the button again
      // cannot hand the successful half over a second time.
      const refused = new Set(result.blocked.map((entry) => entry.line.lineId));
      setBasket((current) => ({
        ...current,
        lines: current.lines.filter((line) => refused.has(line.lineId)),
      }));
    } catch (thrown) {
      setProgress(null);
      setFailure(thrown.message || 'The handover could not be recorded');
    }
  };

  const person = basket.person;
  const count = unitCount(basket);
  const done = report && !report.blocked.length && !basket.lines.length;

  return (
    <AppShell
      title="Hand over"
      subtitle="Give items to somebody, on loan or for keeps"
      actions={person && basket.lines.length > 0 && (
        <Button icon={Check} onClick={handOver} disabled={Boolean(progress)}>
          {progress ? 'Recording…' : `Hand over ${count} item${count === 1 ? '' : 's'}`}
        </Button>
      )}
    >
      {error && <ErrorBanner message={error} onRetry={reload} />}
      {failure && <ErrorBanner message={failure} onRetry={handOver} />}

      {progress && (
        <Card className="as-progress">
          <strong>{PHASE_LABEL[progress.phase] ?? 'Working'}</strong>
          {progress.total > 0 && (
            <div className="bar-track">
              <span
                className="bar-fill"
                style={{ width: `${Math.round((progress.done / progress.total) * 100)}%` }}
              />
            </div>
          )}
        </Card>
      )}

      {report && (
        <Card className={`as-notice ${report.blocked.length ? 'as-notice-warn' : 'as-notice-ok'}`}>
          {report.blocked.length ? <AlertTriangle size={16} /> : <Check size={16} />}
          <span>
            {report.handedOver} item{report.handedOver === 1 ? '' : 's'} handed to {person?.name}.
            {report.blocked.length > 0 && ` ${report.blocked.length} refused — still in the basket below.`}
            {report.staleRows.length > 0
              && ' Some register rows could not be updated; the handover itself is recorded.'}
          </span>
          {done && (
            <span className="as-notice-links">
              <button
                type="button"
                className="as-link"
                onClick={() => navigate(`/assets/people/${encodeURIComponent(person.email)}`)}
              >
                See what {person.name} holds
              </button>
            </span>
          )}
        </Card>
      )}

      <Card className="as-panel">
        <PersonPicker person={person} onChange={(next) => setBasket({ ...basket, person: next })} />
      </Card>

      {person && (
        <>
          <Card className="as-panel">
            <h2 className="as-h2">For how long</h2>
            <div className="as-form">
              <label className="as-field">
                <span className="as-field-label">These are</span>
                <select
                  value={basket.kind}
                  onChange={(event) => setBasket({ ...basket, kind: event.target.value })}
                >
                  <option value={HANDOVER_KIND.ISSUED}>Issued — theirs to keep</option>
                  <option value={HANDOVER_KIND.BORROWED}>Borrowed — coming back</option>
                </select>
              </label>

              {basket.kind === HANDOVER_KIND.BORROWED && (
                <label className="as-field">
                  <span className="as-field-label">Due back</span>
                  <input
                    type="date"
                    value={toLocalDate(basket.dueOn)}
                    onChange={(event) => setBasket({
                      ...basket,
                      // A half-typed date must not wipe the one already set.
                      dueOn: Number.isNaN(Date.parse(event.target.value))
                        ? basket.dueOn
                        : Date.parse(event.target.value),
                    })}
                  />
                </label>
              )}
            </div>
          </Card>

          <Card className="as-panel">
            <h2 className="as-h2">What they are getting</h2>

            <div className="as-addrow">
              <span className="as-searchbox">
                <Search size={15} />
                <input
                  value={query}
                  onChange={(event) => setQuery(event.target.value)}
                  placeholder="Search the register — model, serial, label…"
                />
              </span>
              <Button
                variant={scanning ? 'secondary' : 'primary'}
                size="sm"
                icon={scanning ? X : ScanLine}
                onClick={() => setScanning(!scanning)}
              >
                {scanning ? 'Stop scanning' : 'Scan'}
              </Button>
            </div>

            {matches.length > 0 && (
              <ul className="as-matches">
                {matches.map((asset) => (
                  <li key={asset.id}>
                    <button type="button" className="as-match" onClick={() => add(asset)}>
                      <Plus size={13} />
                      <span className="as-person-text">
                        <strong>{asset.title}</strong>
                        <span className="as-sub">
                          {asset.category}
                          {asset.trackingMode === TRACKED
                            ? ''
                            : ` · ${available(asset)} of ${owned(asset)} available`}
                        </span>
                      </span>
                    </button>
                  </li>
                ))}
              </ul>
            )}

            {query.trim() && !loading && matches.length === 0 && (
              <p className="as-hint">Nothing available matches that.</p>
            )}

            {scanning && (
              <div className="as-viewfinder as-viewfinder-short">
                <video ref={videoRef} playsInline muted className="as-video" />
                <div className="as-reticle" aria-hidden="true" />
                {flash && (
                  <div className={`as-flash as-flash-${flash.kind}`} role="status">
                    {flash.kind === 'dup' ? <AlertTriangle size={15} /> : <Check size={15} />}
                    <span>{flash.text}</span>
                  </div>
                )}
                {state === CAMERA_STATE.DENIED && (
                  <p className="as-camera-msg">
                    This browser is not allowing the camera. Search for the items instead.
                  </p>
                )}
              </div>
            )}

            {basket.lines.length === 0 ? (
              <EmptyState>
                <Package size={20} />
                Nothing in the basket yet. Search for an item, or scan its barcode.
              </EmptyState>
            ) : (
              <ul className="as-basket">
                {basket.lines.map((line) => {
                  const asset = byId.get(line.assetId);
                  const refusal = refusals.get(line.lineId);

                  return (
                    <li key={line.lineId} className={refusal ? 'as-basketline bad' : 'as-basketline'}>
                      <span className="as-person-text">
                        <strong>{line.itemTitle}</strong>
                        <span className="as-sub">{line.category}</span>
                        {refusal && <span className="as-field-issue">{refusal}</span>}
                      </span>

                      {line.trackingMode !== TRACKED && (
                        <label className="as-qtybox">
                          <span className="as-sub">Qty</span>
                          <input
                            type="number"
                            min="1"
                            max={asset ? available(asset) : undefined}
                            inputMode="numeric"
                            value={line.quantity}
                            onChange={(event) => setBasket(
                              setQuantity(basket, line.lineId, event.target.value),
                            )}
                          />
                        </label>
                      )}

                      <select
                        className="as-linekind"
                        value={line.kind ?? basket.kind}
                        onChange={(event) => setBasket(replaceLine(basket, {
                          ...line, kind: event.target.value,
                        }))}
                        aria-label="Issued or borrowed"
                      >
                        <option value={HANDOVER_KIND.ISSUED}>Issued</option>
                        <option value={HANDOVER_KIND.BORROWED}>Borrowed</option>
                      </select>

                      <button
                        type="button"
                        className="as-iconbtn"
                        onClick={() => setBasket(removeLine(basket, line.lineId))}
                        aria-label="Take out of the basket"
                      >
                        <Trash2 size={14} />
                      </button>
                    </li>
                  );
                })}
              </ul>
            )}

            {blockedCount > 0 && (
              <p className="as-field-issue">
                {blockedCount} line{blockedCount === 1 ? '' : 's'} cannot go out. The rest still will.
              </p>
            )}
          </Card>
        </>
      )}
    </AppShell>
  );
}

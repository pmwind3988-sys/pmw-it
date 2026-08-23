import { useCallback, useEffect, useMemo, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card, ErrorBanner, EmptyState } from '../components/ui/Surfaces';
import {
  Save, Plus, Trash2, WifiOff, AlertTriangle, Check, Truck,
} from '../components/ui/Icons';
import { useAssets, SHAREPOINT_SITE_URL } from '../features/assets/useAssets';
import { useSharePointToken } from '../hooks/useRequests';
import { useMsal } from '@azure/msal-react';
import {
  loadBatch, saveBatch, deleteBatch, loadPhoto,
} from '../features/assets/store/assetDb';
import {
  replaceDraft, removeDraft, addDraft, resolveDrafts, batchTitle, BATCH_STATUS,
} from '../features/assets/draft/batch';
import { newDraft, draftIssues } from '../features/assets/draft/draftAsset';
import { indexByTag, normaliseCode } from '../features/assets/identity';
import { saveBatchToSharePoint, remainingDrafts } from '../features/assets/sharepoint/saveBatch';
import DraftCard from '../features/assets/ui/DraftCard';

/**
 * Reviewing a delivery and committing it.
 *
 * This is the one screen where a scanning session becomes a record everybody
 * else can see. Everything guessed is flagged and editable, everything that
 * would collide is blocked by the row rather than by the batch, and the Save
 * button is the only thing on the page that touches the network.
 */

const PHASE_LABEL = {
  provisioning: 'Setting up the SharePoint lists',
  reading: 'Reading what is already in the register',
  photos: 'Uploading photos',
  writing: 'Saving items',
  logging: 'Recording what changed',
  delivery: 'Recording the delivery',
};

export default function AssetBatchPage() {
  const { id } = useParams();
  const navigate = useNavigate();
  const { instance } = useMsal();
  const getToken = useSharePointToken();
  const { assets, loading: loadingRegister, reload } = useAssets();

  const [batch, setBatch] = useState(null);
  const [loading, setLoading] = useState(true);
  const [progress, setProgress] = useState(null);
  const [report, setReport] = useState(null);
  const [error, setError] = useState('');
  const [offline, setOffline] = useState(!navigator.onLine);

  useEffect(() => {
    const update = () => setOffline(!navigator.onLine);
    window.addEventListener('online', update);
    window.addEventListener('offline', update);
    return () => {
      window.removeEventListener('online', update);
      window.removeEventListener('offline', update);
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    loadBatch(id).then((found) => {
      if (!cancelled) {
        setBatch(found ?? null);
        setLoading(false);
      }
    });
    return () => { cancelled = true; };
  }, [id]);

  /**
   * Written back on every edit rather than on a Save-draft button: this batch
   * is the only copy of a delivery somebody has already walked away from the
   * shelf with.
   */
  const update = useCallback((next) => {
    setBatch(next);
    saveBatch(next).catch(() => {});
  }, []);

  const registerTags = useMemo(() => indexByTag(assets), [assets]);

  const resolved = useMemo(() => (batch ? resolveDrafts(batch) : []), [batch]);

  /** Labels claimed earlier in this same batch, so two rows cannot share one. */
  const batchTags = useMemo(() => {
    const map = new Map();
    for (const draft of resolved) {
      const tag = normaliseCode(draft.assetTag);
      if (tag && !map.has(tag)) map.set(tag, draft.localId);
    }
    return map;
  }, [resolved]);

  const issuesByRow = useMemo(() => new Map(
    resolved.map((draft) => [
      draft.localId,
      draftIssues(draft, { registerTags, batchTags }),
    ]),
  ), [resolved, registerTags, batchTags]);

  const blockedCount = [...issuesByRow.values()].filter(
    (issues) => issues.some((issue) => issue.blocking),
  ).length;

  const save = async () => {
    setError('');
    setReport(null);
    try {
      const tokenRes = await getToken();
      const account = instance.getActiveAccount();

      const result = await saveBatchToSharePoint({
        siteUrl: SHAREPOINT_SITE_URL,
        token: tokenRes.accessToken,
        batch,
        photoFor: (photoId) => loadPhoto(photoId),
        savedBy: account?.username ?? account?.name ?? '',
        onProgress: setProgress,
      });

      setReport(result);
      setProgress(null);
      reload();

      const left = remainingDrafts(batch, result);
      if (!left.length) {
        // Nothing is waiting any more, so the batch and its photos go — a
        // saved delivery kept on the phone is storage nothing will ever free.
        await deleteBatch(batch.id);
        setBatch({ ...batch, drafts: [], status: BATCH_STATUS.SAVED });
      } else {
        // Only what did not land stays behind, so pressing Save again does not
        // fill the change log with phantom edits to rows that were fine.
        update({ ...batch, drafts: left });
      }
    } catch (failure) {
      setProgress(null);
      setError(failure.message || 'The delivery could not be saved');
    }
  };

  const discard = async () => {
    await deleteBatch(id);
    navigate('/assets');
  };

  if (loading) {
    return <AppShell title="Delivery"><div className="spinner" /></AppShell>;
  }

  if (!batch) {
    return (
      <AppShell title="Delivery">
        <EmptyState>
          That delivery is not on this device. Batches live on the phone that scanned
          them until they are saved.
        </EmptyState>
      </AppShell>
    );
  }

  const saved = report && !remainingDrafts(batch, report).length;

  return (
    <AppShell
      title={batchTitle(batch)}
      subtitle={`${batch.drafts.length} item${batch.drafts.length === 1 ? '' : 's'} waiting to be saved`}
      actions={(
        <>
          <Button
            variant="secondary"
            icon={Plus}
            onClick={() => update(addDraft(batch, newDraft({ scanSource: 'Manual' })))}
          >
            Add row
          </Button>
          <Button
            icon={Save}
            onClick={save}
            disabled={offline || !batch.drafts.length || Boolean(progress)}
          >
            {progress ? 'Saving…' : 'Save to SharePoint'}
          </Button>
        </>
      )}
    >
      {offline && (
        <Card className="as-notice as-notice-warn">
          <WifiOff size={16} />
          <span>
            You are offline. The delivery is safe on this device — come back when
            you have a connection and press Save.
          </span>
        </Card>
      )}

      {error && <ErrorBanner message={error} onRetry={save} />}

      {progress && (
        <Card className="as-progress">
          <strong>{PHASE_LABEL[progress.phase] ?? 'Working'}</strong>
          {progress.total > 0 && (
            <>
              <div className="bar-track">
                <span
                  className="bar-fill"
                  style={{ width: `${Math.round((progress.done / progress.total) * 100)}%` }}
                />
              </div>
              <span className="as-progress-count">{progress.done} of {progress.total}</span>
            </>
          )}
        </Card>
      )}

      {report && (
        <Card className={`as-notice ${saved ? 'as-notice-ok' : 'as-notice-warn'}`}>
          {saved ? <Check size={16} /> : <AlertTriangle size={16} />}
          <span>
            {report.results.filter((r) => r.action === 'insert' && !r.error).length} added,
            {' '}{report.results.filter((r) => r.action === 'update' && !r.error).length} updated,
            {' '}{report.unchanged} already up to date.
            {report.blocked.length > 0 && ` ${report.blocked.length} refused.`}
            {report.results.some((r) => r.error) && ' Some rows failed and are still here.'}
            {report.photoFailures.length > 0
              && ` ${report.photoFailures.length} photo(s) could not be uploaded; the items saved without them.`}
          </span>
        </Card>
      )}

      {saved && (
        <EmptyState>
          <Truck size={22} />
          <p>This delivery is in the register.</p>
          <Button onClick={() => navigate('/assets')}>Open the register</Button>
        </EmptyState>
      )}

      {!saved && (
        <>
          {blockedCount > 0 && (
            <Card className="as-notice as-notice-bad">
              <AlertTriangle size={16} />
              <span>
                {blockedCount} row{blockedCount === 1 ? '' : 's'} cannot be saved yet —
                fix the label clash below. The rest will still save.
              </span>
            </Card>
          )}

          {loadingRegister && <p className="as-hint">Checking the register for duplicates…</p>}

          <div className="as-review">
            {resolved.map((draft, index) => (
              <DraftCard
                key={draft.localId}
                index={index + 1}
                draft={batch.drafts[index]}
                issues={issuesByRow.get(draft.localId) ?? []}
                onChange={(next) => update(replaceDraft(batch, next))}
                onRemove={() => update(removeDraft(batch, draft.localId))}
              />
            ))}
          </div>

          {!batch.drafts.length && (
            <EmptyState>
              Nothing in this delivery yet. Add a row, or scan some barcodes.
            </EmptyState>
          )}

          <div className="as-actions">
            <Button variant="ghost" icon={Trash2} onClick={discard}>
              Discard this delivery
            </Button>
          </div>
        </>
      )}
    </AppShell>
  );
}

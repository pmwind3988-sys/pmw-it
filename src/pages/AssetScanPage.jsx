import { useCallback, useEffect, useRef, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import AppShell from '../components/AppShell';
import Button from '../components/ui/Button';
import { Card } from '../components/ui/Surfaces';
import {
  Camera, ScanLine, Boxes, Check, X, AlertTriangle, Package,
} from '../components/ui/Icons';
import { useScanner, CAMERA_STATE } from '../features/assets/scan/useScanner';
import {
  createSession, seeCode, commitPool, discardPool, removeDraft, SCAN_MODES, OUTCOMES,
} from '../features/assets/scan/scanSession';
import { signalAccepted, signalDuplicate, signalDone } from '../features/assets/scan/feedback';
import { shrinkToBlob } from '../features/assets/scan/capturePhoto';
import { newBatch } from '../features/assets/draft/batch';
import { newId } from '../features/assets/draft/draftAsset';
import { saveBatch, savePhoto } from '../features/assets/store/assetDb';
import PhotoInput from '../features/assets/ui/PhotoInput';
import ScanControls from '../features/assets/ui/ScanControls';
import { useScrollLock } from '../hooks/useScrollLock';

/**
 * Scanning a delivery.
 *
 * Three steps, in the order the job actually happens: the paperwork (typed
 * once for the whole delivery), the mode, then the camera. Nothing here
 * touches SharePoint — the batch is written to this device and reviewed later,
 * which is what makes a store room with no signal a non-issue (§4.1).
 */

const STEP = { PURCHASE: 'purchase', MODE: 'mode', SCANNING: 'scanning' };

/** `datetime-local` wants the local wall clock, not an ISO instant. */
function toLocalInput(epochMs) {
  const date = new Date(epochMs);
  const offset = date.getTimezoneOffset() * 60000;
  return new Date(epochMs - offset).toISOString().slice(0, 16);
}

export default function AssetScanPage() {
  const navigate = useNavigate();
  const [step, setStep] = useState(STEP.PURCHASE);
  const [batch, setBatch] = useState(() => newBatch());
  const [session, setSession] = useState(() => createSession(SCAN_MODES.MANY));
  const [flash, setFlash] = useState(null);
  const flashTimer = useRef(null);

  const scanning = step === STEP.SCANNING;

  // The camera covers the screen while it is open, so the page behind it must
  // not scroll under the picture.
  useScrollLock(scanning);

  const show = useCallback((kind, text) => {
    setFlash({ kind, text });
    if (flashTimer.current) clearTimeout(flashTimer.current);
    flashTimer.current = setTimeout(() => setFlash(null), 1400);
  }, []);

  useEffect(() => () => {
    if (flashTimer.current) clearTimeout(flashTimer.current);
  }, []);

  /**
   * Every code the camera reads. The session decides what each one means; this
   * only reports the answer, because the person holding the phone is looking
   * at a box rather than at the screen (§4.4).
   */
  const onCodes = useCallback((codes) => {
    setSession((current) => {
      let next = current;
      for (const code of codes) {
        const step2 = seeCode(next, code);
        next = step2.session;

        if (step2.outcome === OUTCOMES.DUPLICATE) {
          signalDuplicate();
          show('dup', `Already scanned — ${step2.code}`);
        } else if (step2.outcome === OUTCOMES.SHARED) {
          // Said out loud, because it looks like the scanner accepting a
          // duplicate. It is the opposite: the code both boxes carry is the
          // one thing that proves it is not the serial.
          signalAccepted();
          show('ok', `Same as an earlier box — part number ${step2.code}`);
        } else if (step2.outcome !== OUTCOMES.EMPTY) {
          signalAccepted();
          show('ok', step2.code);
        }
      }
      return next;
    });
  }, [show]);

  const {
    videoRef, state, grabFrame, usingPonyfill,
    controls, torchOn, toggleTorch, zoomTo, focusOn, quiet,
  } = useScanner({ active: scanning, onCodes });

  /** A tap on the picture is "focus there" — the coordinates the camera wants
   *  are the tap position as a fraction of the element it landed on. */
  const tapToFocus = (event) => {
    const box = event.currentTarget.getBoundingClientRect();
    if (!box.width || !box.height) return;
    focusOn((event.clientX - box.left) / box.width, (event.clientY - box.top) / box.height);
  };

  const setPurchase = (field) => (event) => setBatch((current) => ({
    ...current,
    purchase: { ...current.purchase, [field]: event.target.value },
  }));

  const start = (mode) => {
    setSession(createSession(mode));
    setStep(STEP.SCANNING);
  };

  /** ONE mode: the box in frame is done. */
  const confirmBox = async () => {
    const photoId = await capture();
    setSession((current) => commitPool(current, photoId ? { photoId } : {}).session);
    signalDone();
    show('ok', 'Item added');
  };

  /** A still off the live camera, so the item is photographed where it stands. */
  const capture = async () => {
    const frame = grabFrame();
    if (!frame?.videoWidth) return null;
    try {
      const blob = await shrinkToBlob(frame);
      const id = newId();
      await savePhoto(id, blob);
      return id;
    } catch {
      return null;
    }
  };

  const finish = async () => {
    const finished = { ...batch, drafts: session.drafts };
    await saveBatch(finished);
    navigate(`/assets/batch/${finished.id}`);
  };

  const count = session.drafts.length;

  return (
    <AppShell
      title="Scan a delivery"
      subtitle={
        scanning
          ? 'Point the camera at the barcodes. Nothing is sent anywhere until you review it.'
          : 'The purchase details are typed once and copied onto every item in this delivery.'
      }
      actions={scanning && (
        <Button variant="secondary" icon={X} onClick={() => setStep(STEP.MODE)}>
          Stop
        </Button>
      )}
    >
      {step === STEP.PURCHASE && (
        <Card className="as-panel">
          <h2 className="as-h2">Where this delivery came from</h2>
          <p className="as-hint">All of it is optional, and all of it can be changed later.</p>

          <div className="as-form">
            <label className="as-field">
              <span className="as-field-label">Purchased from</span>
              <input
                value={batch.purchase.supplier}
                onChange={setPurchase('supplier')}
                placeholder="Ingram Micro"
              />
            </label>

            <label className="as-field">
              <span className="as-field-label">PO number</span>
              <input
                value={batch.purchase.poNumber}
                onChange={setPurchase('poNumber')}
                placeholder="PO-4471"
              />
            </label>

            <label className="as-field">
              <span className="as-field-label">DO number</span>
              <input
                value={batch.purchase.doNumber}
                onChange={setPurchase('doNumber')}
                placeholder="DO-8891"
              />
            </label>

            <label className="as-field">
              <span className="as-field-label">Arrived on</span>
              <input
                type="datetime-local"
                value={toLocalInput(batch.purchase.arrivedOn)}
                onChange={(event) => setBatch((current) => ({
                  ...current,
                  purchase: {
                    ...current.purchase,
                    // An unparseable half-typed date must not wipe the value
                    // the field already had.
                    arrivedOn: Number.isNaN(Date.parse(event.target.value))
                      ? current.purchase.arrivedOn
                      : Date.parse(event.target.value),
                  },
                }))}
              />
            </label>

            <label className="as-field">
              <span className="as-field-label">Remarks</span>
              <textarea rows={2} value={batch.purchase.remarks} onChange={setPurchase('remarks')} />
            </label>
          </div>

          {/* The delivery that arrived before there was anywhere to record it.
              A register that will not take it until every serial has been
              found is a register that never gets it at all, so it is taken as
              it stands and MARKED -- one switch rather than a question against
              each of thirty blanks (`detailsPending.js`). */}
          <label className="as-switch">
            <input
              type="checkbox"
              checked={Boolean(batch.purchase.detailsPending)}
              onChange={(event) => setBatch((current) => ({
                ...current,
                purchase: { ...current.purchase, detailsPending: event.target.checked },
              }))}
            />
            <span>
              <strong>An older delivery I am entering late</strong>
              <span className="as-hint">
                The DO number, the serial numbers or the photos are missing. Nothing
                will nag you about them, every item is marked as needing details,
                and you can finish them off later from the register.
              </span>
            </span>
          </label>

          <div className="as-po-photo">
            <span className="as-field-label">Photo of the PO or delivery note</span>
            <PhotoInput
              photoId={batch.purchase.poPhotoId}
              onChange={(poPhotoId) => setBatch((current) => ({
                ...current,
                purchase: { ...current.purchase, poPhotoId },
              }))}
              label="Snap the paperwork"
            />
          </div>

          <div className="as-actions">
            <Button onClick={() => setStep(STEP.MODE)}>Next</Button>
            <Button variant="ghost" onClick={() => navigate('/assets')}>Cancel</Button>
          </div>
        </Card>
      )}

      {step === STEP.MODE && (
        <div className="as-modes">
          <button type="button" className="as-mode" onClick={() => start(SCAN_MODES.MANY)}>
            <Boxes size={26} />
            <strong>Many items</strong>
            <span>
              Sweep along a shelf. Every new barcode becomes its own item — fastest
              when each box shows one code.
            </span>
          </button>

          <button type="button" className="as-mode" onClick={() => start(SCAN_MODES.ONE)}>
            <Package size={26} />
            <strong>One item at a time</strong>
            <span>
              For a box with several barcodes on it. Read them all, then confirm —
              the serial and the part number are worked out together.
            </span>
          </button>

          {count > 0 && (
            <div className="as-actions">
              <Button icon={Check} onClick={finish}>
                Review {count} item{count === 1 ? '' : 's'}
              </Button>
            </div>
          )}
        </div>
      )}

      {scanning && (
        <div className="as-scan as-fullcam">
          <div className="as-viewfinder" onClick={tapToFocus}>
            <video ref={videoRef} playsInline muted className="as-video" />
            <div className="as-reticle" aria-hidden="true" />

            <ScanControls
              controls={controls}
              torchOn={torchOn}
              onTorch={toggleTorch}
              onZoom={zoomTo}
            />

            {/* Said only after a while, and only while the camera is
                otherwise fine: the two things that fix an unread barcode are
                filling the box with it and turning a light on, and neither is
                obvious from a viewfinder that simply does nothing. */}
            {quiet && state === CAMERA_STATE.RUNNING && (
              <p className="as-camera-hint">
                Nothing read yet. Fill the white box with the barcode, move a little
                closer{controls.torch ? ', or turn the light on' : ''} — and tap the
                picture to focus.
              </p>
            )}

            {flash && (
              <div className={`as-flash as-flash-${flash.kind}`} role="status">
                {flash.kind === 'dup' ? <AlertTriangle size={15} /> : <Check size={15} />}
                <span>{flash.text}</span>
              </div>
            )}

            {state === CAMERA_STATE.STARTING && (
              <p className="as-camera-msg">Starting the camera…</p>
            )}

            {state === CAMERA_STATE.DENIED && (
              <div className="as-camera-msg">
                <Camera size={22} />
                <p>
                  This browser is not allowing the camera. Turn it on for this site in
                  your browser settings, or add the items by hand.
                </p>
                <Button variant="secondary" onClick={finish}>Add by hand instead</Button>
              </div>
            )}

            {(state === CAMERA_STATE.UNAVAILABLE || state === CAMERA_STATE.NO_DECODER) && (
              <div className="as-camera-msg">
                <ScanLine size={22} />
                <p>
                  {state === CAMERA_STATE.NO_DECODER
                    ? 'This browser cannot read barcodes. Chrome on Android or Safari on a recent iPhone can.'
                    : 'No camera is available on this device.'}
                </p>
                <Button variant="secondary" onClick={finish}>Add by hand instead</Button>
              </div>
            )}
          </div>

          <div className="as-scanbar">
            {session.mode === SCAN_MODES.ONE ? (
              <>
                <span className="as-pool">
                  {session.pool.length
                    ? `${session.pool.length} code${session.pool.length === 1 ? '' : 's'} on this box`
                    : 'Read every barcode on the box'}
                </span>
                <Button
                  icon={Check}
                  onClick={confirmBox}
                  disabled={!session.pool.length}
                >
                  This box is done
                </Button>
                {session.pool.length > 0 && (
                  <Button
                    variant="ghost"
                    onClick={() => setSession((current) => discardPool(current))}
                  >
                    Discard
                  </Button>
                )}
              </>
            ) : (
              <span className="as-pool">
                {count ? `${count} item${count === 1 ? '' : 's'} scanned` : 'Waiting for a barcode'}
              </span>
            )}

            <Button variant="primary" icon={Check} onClick={finish} disabled={!count}>
              Done ({count})
            </Button>
          </div>

          {usingPonyfill && (
            <p className="as-hint as-hint-inline">
              This browser has no built-in barcode reader, so scanning is a little slower.
            </p>
          )}

          <ul className="as-strip">
            {session.drafts.map((draft, index) => (
              <li key={draft.localId}>
                <span className="as-strip-index">{index + 1}</span>
                <span className="as-strip-code">
                  {draft.serialNumber || draft.partNumber || draft.assetTag || 'code'}
                </span>
                <button
                  type="button"
                  className="as-iconbtn"
                  onClick={() => setSession((current) => removeDraft(current, draft.localId))}
                  aria-label="Remove"
                >
                  <X size={12} />
                </button>
              </li>
            ))}
          </ul>
        </div>
      )}
    </AppShell>
  );
}

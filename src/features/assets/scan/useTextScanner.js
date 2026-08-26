import { useCallback, useEffect, useRef, useState } from 'react';
import { CAMERA_STATE, useCamera } from './useCamera.js';
import { createTextReader, readTextFrame, READER_SOURCE } from './textReader.js';
import { readTextFields } from './classifyText.js';
import {
  newTextScan, recordReading, rejectValue, dismissExtra, MAX_PASSES,
} from './textScan.js';

/**
 * Pointing the camera at a label until the words stop changing.
 *
 * The loop is sequential rather than timed: recognising one frame takes
 * the best part of a second, so a fixed interval would only queue frames
 * behind an engine that is already busy. Read, wait for the answer, read
 * again.
 *
 * What each pass MEANS lives in `classifyText.js` and `textScan.js`,
 * which are pure and tested. This holds only the parts that cannot be:
 * the camera, the engine, and the permissions the browser may refuse.
 *
 * The hook keeps no state that has to be reset. Whoever opens the camera
 * mounts it and whoever closes it unmounts it, so a second scan starts
 * from nothing without anything here having to remember to clear.
 */

export const SCAN_STATE = {
  STARTING: 'starting',
  READING: 'reading',
  DONE: 'done',
  DENIED: 'denied',
  UNAVAILABLE: 'unavailable',
  NO_READER: 'no-reader',
};

export function useTextScanner({ active = true } = {}) {
  const [scan, setScan] = useState(newTextScan);
  const [reader, setReader] = useState(null);
  const [done, setDone] = useState(false);
  const scanRef = useRef(scan);
  const doneRef = useRef(false);

  const camera = useCamera({ active });

  useEffect(() => {
    scanRef.current = scan;
  }, [scan]);

  // The engine is loaded once and thrown away with the sheet. Keeping a
  // worker alive behind a closed camera holds several hundred megabytes
  // for a scan that may never come.
  useEffect(() => {
    if (!active) return undefined;

    let cancelled = false;
    let loaded = null;

    (async () => {
      const created = await createTextReader();
      loaded = created;
      if (cancelled) {
        created.terminate();
        return;
      }
      setReader(created);
    })();

    return () => {
      cancelled = true;
      loaded?.terminate();
    };
  }, [active]);

  const ready = camera.state === CAMERA_STATE.RUNNING && Boolean(reader?.read);

  useEffect(() => {
    if (!ready) return undefined;

    let cancelled = false;

    const pass = async () => {
      if (cancelled || doneRef.current) return;

      const frame = camera.videoRef.current;
      // A frame with no dimensions yet reads as nothing and costs a whole
      // pass, so it is skipped rather than recognised.
      if (frame?.videoWidth) {
        const lines = await readTextFrame(reader.read, frame);
        if (cancelled) return;

        if (lines.length) {
          const next = recordReading(scanRef.current, readTextFields(lines));
          scanRef.current = next;
          setScan(next);

          // It keeps reading after a value settles rather than stopping on
          // the first complete answer. Nothing is written into the form any
          // more, so another pass costs only battery -- and while somebody is
          // reading the list, the rest of the label is still worth picking
          // up. `MAX_PASSES` is what ends it.
          if (next.passes >= MAX_PASSES) {
            doneRef.current = true;
            setDone(true);
            return;
          }
        }
      }

      if (!cancelled) pass();
    };

    pass();

    return () => { cancelled = true; };
  }, [ready, reader, camera.videoRef]);

  /** Stop early and keep whatever has settled so far. */
  const finish = useCallback(() => {
    doneRef.current = true;
    setDone(true);
  }, []);

  /**
   * Crossed out. Written through the ref as well as through state because the
   * reading loop reads the ref, and a refusal that only reached state would be
   * overwritten by the pass already in flight.
   */
  const reject = useCallback((field) => {
    const next = rejectValue(scanRef.current, field);
    scanRef.current = next;
    setScan(next);
  }, []);

  const dismiss = useCallback((value) => {
    const next = dismissExtra(scanRef.current, value);
    scanRef.current = next;
    setScan(next);
  }, []);

  let state = SCAN_STATE.STARTING;
  if (reader && !reader.read) state = SCAN_STATE.NO_READER;
  else if (camera.state === CAMERA_STATE.DENIED) state = SCAN_STATE.DENIED;
  else if (camera.state === CAMERA_STATE.UNAVAILABLE) state = SCAN_STATE.UNAVAILABLE;
  else if (done) state = SCAN_STATE.DONE;
  else if (ready) state = SCAN_STATE.READING;

  return {
    videoRef: camera.videoRef,
    state,
    error: camera.error,
    /** Which engine is doing the reading, for anything that wants to say so. */
    source: reader?.source ?? READER_SOURCE.NONE,
    scan,
    finish,
    reject,
    dismiss,
  };
}

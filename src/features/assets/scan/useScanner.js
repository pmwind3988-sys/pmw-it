import { useCallback, useEffect, useRef, useState } from 'react';
import { createDetector, readFrame, DETECTOR_SOURCE } from './detector';
import { cropToCanvas } from './cropRegion';
import {
  VIDEO_CONSTRAINTS, trackControls, setTorch, setZoom, focusAt,
} from './cameraTrack';

/**
 * The camera, running.
 *
 * Everything that decides what a code MEANS lives in `scanSession.js`, which
 * is pure and tested. This holds only the parts that cannot be: the video
 * stream, the decode loop, and the permissions the browser may refuse.
 *
 * `onCodes(codes)` is called with every frame's worth of decoded values. It is
 * held in a ref rather than being a dependency of the loop, because a handler
 * that closes over the session changes identity on every scan — and a loop
 * that restarts on every scan drops frames exactly when it is busiest.
 */

export const CAMERA_STATE = {
  STARTING: 'starting',
  RUNNING: 'running',
  DENIED: 'denied',
  UNAVAILABLE: 'unavailable',
  NO_DECODER: 'no-decoder',
};

/**
 * How long to wait before admitting nothing is being read.
 *
 * Long enough not to nag somebody still lining the box up, short enough to
 * arrive while they are still holding the phone over it rather than after they
 * have given up and put it down.
 */
export const QUIET_MS = 8000;

/** A breath between passes, so a decoder that answers instantly does not
 *  monopolise the main thread and stall the video it is reading. */
const BREATH_MS = 30;

export function useScanner({ active = true, onCodes } = {}) {
  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const handlerRef = useRef(onCodes);
  const [state, setState] = useState(CAMERA_STATE.STARTING);
  const [detectorSource, setDetectorSource] = useState(null);
  const [error, setError] = useState('');
  const [controls, setControls] = useState({ torch: false, zoom: null });
  const [torchOn, setTorchOn] = useState(false);
  const [quiet, setQuiet] = useState(false);

  // In an effect rather than during render: a ref written while rendering is
  // a value React may have already read, and eslint fails the build over it.
  useEffect(() => {
    handlerRef.current = onCodes;
  }, [onCodes]);

  useEffect(() => {
    if (!active) return undefined;

    let cancelled = false;
    let timer = null;

    const stop = () => {
      if (timer) clearTimeout(timer);
      for (const track of streamRef.current?.getTracks() ?? []) track.stop();
      streamRef.current = null;
    };

    (async () => {
      const detector = await createDetector();
      if (cancelled) return;

      setDetectorSource(detector.source);
      if (!detector.detect) {
        // Nothing on this browser can decode a barcode. Say so and offer the
        // manual path rather than showing a camera that finds nothing forever.
        setState(CAMERA_STATE.NO_DECODER);
        return;
      }

      if (!navigator.mediaDevices?.getUserMedia) {
        setState(CAMERA_STATE.UNAVAILABLE);
        return;
      }

      try {
        // `environment` is the back camera. `ideal` rather than `exact` so a
        // laptop with only a front camera still works instead of throwing.
        streamRef.current = await navigator.mediaDevices.getUserMedia({
          video: VIDEO_CONSTRAINTS,
          audio: false,
        });
      } catch (failure) {
        if (cancelled) return;
        setError(failure?.message ?? '');
        setState(failure?.name === 'NotAllowedError'
          ? CAMERA_STATE.DENIED
          : CAMERA_STATE.UNAVAILABLE);
        return;
      }

      if (cancelled) {
        stop();
        return;
      }

      const video = videoRef.current;
      if (video) {
        video.srcObject = streamRef.current;
        // iOS refuses to play an inline video without both of these, and a
        // video that never plays decodes nothing.
        video.setAttribute('playsinline', 'true');
        video.muted = true;
        await video.play().catch(() => {});
      }

      setState(CAMERA_STATE.RUNNING);
      // What this handset can actually do, asked once the stream exists —
      // capabilities are a property of the running track, not of the request.
      setControls(trackControls(streamRef.current));

      let lastFound = Date.now();
      // Alternated rather than combined: the crop is where the barcode being
      // aimed at is, and the whole frame is the safety net for one just
      // outside the box. Reading both every pass would halve the rate.
      let readCrop = true;

      const tick = async () => {
        if (cancelled) return;

        const frame = videoRef.current;
        // A frame with no dimensions yet decodes to nothing and costs a whole
        // pass, so it is skipped rather than read.
        if (frame?.videoWidth) {
          // The aiming box at the camera's own resolution. A sticker held at
          // arm's length is a handful of pixels wide in the whole frame and
          // resolves to nothing; this is the single biggest reason a barcode
          // is never read (`cropRegion.js`).
          const source = (readCrop && cropToCanvas(frame)) || frame;
          readCrop = !readCrop;

          const codes = await readFrame(detector.detect, source);
          if (!cancelled && codes.length) {
            lastFound = Date.now();
            setQuiet(false);
            handlerRef.current?.(codes);
          } else if (!cancelled && Date.now() - lastFound > QUIET_MS) {
            setQuiet(true);
          }
        }

        // Sequential, not on a fixed interval. A software decoder on an iPhone
        // takes longer than any interval worth setting, so timed passes queue
        // up behind an engine that is already busy and the scan gets slower
        // the harder it is working. Read, wait for the answer, read again.
        if (!cancelled) timer = setTimeout(tick, BREATH_MS);
      };

      timer = setTimeout(tick, BREATH_MS);
    })();

    return () => {
      cancelled = true;
      stop();
    };
  }, [active]);

  /** A still from the live camera, for the photo of the item being scanned. */
  const grabFrame = useCallback(() => videoRef.current ?? null, []);

  const toggleTorch = useCallback(async () => {
    const wanted = !torchOn;
    // Only believed once the camera says yes: a button that latches on a
    // handset that refused the torch is a light that is never coming on.
    if (await setTorch(streamRef.current, wanted)) setTorchOn(wanted);
  }, [torchOn]);

  const zoomTo = useCallback((value) => setZoom(streamRef.current, value), []);

  /** Where the person tapped, in the 0-to-1 coordinates the camera wants. */
  const focusOn = useCallback((x, y) => focusAt(streamRef.current, x, y), []);

  return {
    videoRef,
    state,
    error,
    grabFrame,
    controls,
    torchOn,
    toggleTorch,
    zoomTo,
    focusOn,
    /** Nothing has decoded for a while, so the screen can say what to try. */
    quiet,
    usingPonyfill: detectorSource === DETECTOR_SOURCE.PONYFILL,
  };
}

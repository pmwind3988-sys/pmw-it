import { useCallback, useEffect, useRef, useState } from 'react';
import { createDetector, readFrame, DETECTOR_SOURCE } from './detector';

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

/** ~8 reads a second. Faster wins nothing: the decode is the slow part. */
const FRAME_INTERVAL_MS = 120;

export function useScanner({ active = true, onCodes } = {}) {
  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const handlerRef = useRef(onCodes);
  const [state, setState] = useState(CAMERA_STATE.STARTING);
  const [detectorSource, setDetectorSource] = useState(null);
  const [error, setError] = useState('');

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
          video: {
            facingMode: { ideal: 'environment' },
            width: { ideal: 1280 },
            height: { ideal: 720 },
          },
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

      const tick = async () => {
        if (cancelled) return;

        const frame = videoRef.current;
        // A frame with no dimensions yet decodes to nothing and costs a whole
        // interval, so it is skipped rather than read.
        if (frame?.videoWidth) {
          const codes = await readFrame(detector.detect, frame);
          if (!cancelled && codes.length) handlerRef.current?.(codes);
        }

        if (!cancelled) timer = setTimeout(tick, FRAME_INTERVAL_MS);
      };

      timer = setTimeout(tick, FRAME_INTERVAL_MS);
    })();

    return () => {
      cancelled = true;
      stop();
    };
  }, [active]);

  /** A still from the live camera, for the photo of the item being scanned. */
  const grabFrame = useCallback(() => videoRef.current ?? null, []);

  return {
    videoRef,
    state,
    error,
    grabFrame,
    usingPonyfill: detectorSource === DETECTOR_SOURCE.PONYFILL,
  };
}

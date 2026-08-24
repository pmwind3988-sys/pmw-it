import { useCallback, useEffect, useRef, useState } from 'react';

/**
 * The back camera, running, and nothing else.
 *
 * Split out of the scanning hooks because two of them now need it and
 * what they do with the frames has nothing in common: the barcode
 * scanner reads eight times a second and never waits, while the text
 * scanner reads one frame, waits about a second for the answer, and
 * reads again.
 *
 * `useScanner` keeps its own copy of this for now. It cannot be
 * re-tested without a camera in the room, and a silent regression there
 * would break the flow the whole register is filled through.
 */

export const CAMERA_STATE = {
  STARTING: 'starting',
  RUNNING: 'running',
  DENIED: 'denied',
  UNAVAILABLE: 'unavailable',
};

export function useCamera({ active = true } = {}) {
  const videoRef = useRef(null);
  const streamRef = useRef(null);
  const [state, setState] = useState(CAMERA_STATE.STARTING);
  const [error, setError] = useState('');

  useEffect(() => {
    if (!active) return undefined;

    let cancelled = false;

    const stop = () => {
      for (const track of streamRef.current?.getTracks() ?? []) track.stop();
      streamRef.current = null;
    };

    (async () => {
      if (!navigator.mediaDevices?.getUserMedia) {
        setState(CAMERA_STATE.UNAVAILABLE);
        return;
      }

      try {
        // `environment` is the back camera. `ideal` rather than `exact` so
        // a laptop with only a front camera still works instead of throwing.
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
        // video that never plays is a black rectangle with a camera light on.
        video.setAttribute('playsinline', 'true');
        video.muted = true;
        await video.play().catch(() => {});
      }

      if (!cancelled) setState(CAMERA_STATE.RUNNING);
    })();

    return () => {
      cancelled = true;
      stop();
    };
  }, [active]);

  /** The live element, for whoever wants a frame off it. */
  const grabFrame = useCallback(() => videoRef.current ?? null, []);

  return { videoRef, state, error, grabFrame };
}

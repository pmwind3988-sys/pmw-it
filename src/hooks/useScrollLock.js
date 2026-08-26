import { useEffect } from 'react';

/**
 * Holds the page still behind a full-screen camera.
 *
 * Without it the page underneath scrolls when somebody swipes at the picture,
 * and closing the camera leaves them somewhere they did not choose to be. It
 * also stops the pull-to-refresh gesture on Android throwing away a scanning
 * session mid-delivery, which is the expensive version of the same bug.
 *
 * The previous value is restored rather than being set to a default, because
 * something else on the page may have had an opinion about it first.
 */
export function useScrollLock(active = true) {
  useEffect(() => {
    if (!active) return undefined;

    const { body } = document;
    const previous = body.style.overflow;
    const previousOverscroll = body.style.overscrollBehavior;

    body.style.overflow = 'hidden';
    body.style.overscrollBehavior = 'none';

    return () => {
      body.style.overflow = previous;
      body.style.overscrollBehavior = previousOverscroll;
    };
  }, [active]);
}

import * as React from 'react';

/**
 * True while the component is mounted. Async work checks it before it sets
 * state, so an answer that arrives after the page has moved on is dropped.
 */
export function useMountedRef(): React.MutableRefObject<boolean> {
  const mounted = React.useRef(true);
  React.useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
    };
  }, []);
  return mounted;
}

let instanceCount = 0;

/** A unique id per mounted instance, for aria links (React 17 has no useId). */
export function useInstanceId(prefix: string): string {
  const [id] = React.useState(() => {
    instanceCount += 1;
    return `${prefix}-${instanceCount}`;
  });
  return id;
}

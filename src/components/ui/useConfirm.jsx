import { useCallback, useRef, useState } from 'react';
import ConfirmDialog from './ConfirmDialog';

/**
 * `const { ask, dialog } = useConfirm()`, then `if (!await ask({...})) return;`
 * and render `{dialog}` on the page.
 *
 * A promise rather than a callback, so the thing being guarded reads top to
 * bottom exactly as it did with `window.confirm` — the guard stays one line at
 * the head of the function it guards, where it can be seen.
 */
export function useConfirm() {
  const [request, setRequest] = useState(null);
  // The waiting promise is held in a ref rather than in the state, because
  // React may run a state updater twice and settling the same question twice
  // is the kind of thing that goes unnoticed until it removes something.
  const waiting = useRef(null);

  const ask = useCallback((options) => new Promise((resolve) => {
    waiting.current = resolve;
    setRequest(options);
  }), []);

  const answer = useCallback((said) => {
    const resolve = waiting.current;
    waiting.current = null;
    setRequest(null);
    resolve?.(said);
  }, []);

  const dialog = request
    ? <ConfirmDialog {...request} onAnswer={answer} />
    : null;

  return { ask, dialog };
}

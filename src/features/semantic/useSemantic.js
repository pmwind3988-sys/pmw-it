import { useContext } from 'react';
import { SemanticContext } from './semanticStore.js';

/**
 * In its own file, not beside the provider: a module that exports a
 * component must export nothing else, or it drops out of Fast Refresh
 * and fails `npm run lint`. Same rule that put `initialsOf` in
 * `src/utils/initials.js`.
 */
export function useSemantic() {
  const value = useContext(SemanticContext);
  if (!value) throw new Error('useSemantic must be used inside <SemanticProvider>.');
  return value;
}

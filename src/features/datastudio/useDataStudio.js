import { useContext } from 'react';
import { DataStudioContext } from './dataStudioStore.js';

/**
 * In its own file, not beside the provider: a module that exports a
 * component must export nothing else, or it drops out of Fast Refresh
 * and fails `npm run lint`. Same rule that put `initialsOf` in
 * `src/utils/initials.js`.
 */
export function useDataStudio() {
  const value = useContext(DataStudioContext);
  if (!value) throw new Error('useDataStudio must be used inside <DataStudioProvider>.');
  return value;
}

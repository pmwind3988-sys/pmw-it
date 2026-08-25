import { officeLicense } from './officeLicense.js';
import { gpuClass } from './gpuClass.js';
import { serverDependency } from './serverDependency.js';
import { deviceFit } from './deviceFit.js';

/**
 * The persona layer, laid over a record the moment it is read.
 *
 * Nothing here is stored in SharePoint. Every ingredient it needs — the
 * department, the Office products, the graphics adapters, the mapped drives,
 * the network line — is already on the row, so the verdict is recomputed on the
 * way out rather than written down and left to go stale. Change the memory
 * floor for Engineering tomorrow and every machine re-grades itself on the next
 * page load, with no re-scan and no migration.
 */
export function enrichFit(record) {
  if (!record) return record;

  const withFacts = {
    ...record,
    ...officeLicense(record.microsoftOffice ?? []),
    ...gpuClass(record.gpuList ?? []),
    ...serverDependency(record),
  };

  return { ...withFacts, ...deviceFit(withFacts) };
}

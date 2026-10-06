import {
  GRADES, UNKNOWN, PROFILE_KEYS, STORAGE_TYPES, GRAPHICS_KINDS, WINDOWS_KINDS,
} from './defaultStandard.js';

const CUT_PARTS = ['cpu', 'ram', 'storageSize'];
const ORDER = 'Must run in order: Critical below ≤ Needs attention below ≤ Optimal from.';

/** Whether a standard can be graded with, and every reason it cannot, each with where. */
export function validateStandard(standard) {
  const errors = [];
  const add = (path, message) => errors.push({ path, message });

  if (!standard || typeof standard !== 'object') {
    return { ok: false, errors: [{ path: '', message: 'The standard is empty.' }] };
  }
  if (standard.schema !== 1) add('schema', 'This standard was saved by a newer version of the portal.');

  for (const key of PROFILE_KEYS) {
    const profile = standard.profiles?.[key];
    if (!profile) {
      add(`profiles.${key}`, 'This profile is missing.');
      continue;
    }
    for (const part of CUT_PARTS) {
      const c = profile[part];
      const nums = [c?.criticalBelow, c?.attentionBelow, c?.optimalFrom];
      if (!nums.every((n) => typeof n === 'number' && Number.isFinite(n) && n >= 0)) {
        add(`profiles.${key}.${part}`, 'Every cut-off needs a number of 0 or more.');
      } else if (!(nums[0] <= nums[1] && nums[1] <= nums[2])) {
        add(`profiles.${key}.${part}`, ORDER);
      }
    }
    const choices = [['storageType', STORAGE_TYPES], ['graphics', GRAPHICS_KINDS], ['windows', WINDOWS_KINDS]];
    for (const [part, kinds] of choices) {
      for (const kind of kinds) {
        if (!GRADES.includes(profile[part]?.[kind])) add(`profiles.${key}.${part}.${kind}`, 'Pick a grade.');
      }
    }
  }

  for (const [department, profileKey] of Object.entries(standard.departments ?? {})) {
    if (!PROFILE_KEYS.includes(profileKey)) add(`departments.${department}`, 'Unknown profile.');
  }

  for (const grade of [...GRADES, UNKNOWN]) {
    if (!/^#[0-9a-f]{6}$/i.test(standard.colors?.[grade] ?? '')) add(`colors.${grade}`, 'Use a colour like #1a88de.');
  }

  return { ok: errors.length === 0, errors };
}

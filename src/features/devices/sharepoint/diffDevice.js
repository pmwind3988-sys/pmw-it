import { TRACKED_FIELDS } from './deviceSchema.js';

/**
 * Compared as strings so a number that came back from SharePoint as "8" does
 * not read as a change against the number 8. Null and empty are the same thing
 * here — both mean "we do not have this".
 */
const asText = (value) => (value === null || value === undefined ? '' : String(value));

export function diffDevice(existing, incoming) {
  const changes = [];

  for (const fieldName of TRACKED_FIELDS) {
    const oldValue = asText(existing?.[fieldName]);
    const newValue = asText(incoming?.[fieldName]);

    if (oldValue === newValue) continue;

    let changeType = 'Updated';
    if (!oldValue) changeType = 'Added';
    else if (!newValue) changeType = 'Removed';

    changes.push({ fieldName, oldValue, newValue, changeType });
  }

  return changes;
}

export function indexByName(records) {
  const index = new Map();
  for (const record of records) {
    if (!record.computerName) continue;
    index.set(String(record.computerName).toLowerCase(), record);
  }
  return index;
}

import { spFetch, listPath, ITEM_ACCEPT } from '../../sharepoint/spClient.js';
import { provisionSchema } from '../../sharepoint/provision.js';
import { withRetry } from '../../sharepoint/writePool.js';
import {
  SETTINGS_LIST_NAME, EMAIL_SETTING_KEY, normaliseEmailSettings, serialiseEmailSettings,
} from '../links/emailSettings.js';
import { provisionLinks } from './checklistLinks.js';

/**
 * Reading and saving who is emailed when a checklist is signed, as the
 * signed-in person over SharePoint REST. One row per setting in
 * `IT Portal Settings`: `Title` is the key, `SettingValue` the JSON.
 *
 * SharePoint decides who may save: whoever it lets edit that list. The
 * server only ever reads it.
 */

export const SETTINGS_COLUMNS = [
  { StaticName: 'SettingValue', Title: 'Value', kind: 'note' },
];

async function findRow(siteUrl, token, key) {
  const filter = encodeURIComponent(`Title eq '${key}'`);
  const response = await spFetch(
    siteUrl,
    `${listPath(SETTINGS_LIST_NAME)}/items?$filter=${filter}&$top=1`,
    { token, accept: ITEM_ACCEPT },
  );
  // Nobody has saved a setting yet: the list does not exist, and that is "off".
  if (response.status === 404) return null;
  if (!response.ok) {
    throw new Error(`Could not read the email settings (${response.status}): ${await response.text()}`);
  }
  return (await response.json()).value?.[0] ?? null;
}

export async function readEmailSettings({ siteUrl, token }) {
  const row = await findRow(siteUrl, token, EMAIL_SETTING_KEY);
  return normaliseEmailSettings(row?.SettingValue ?? null);
}

/**
 * Saved as given, after normalising. The links list is provisioned too: the
 * server records each email on the link row, in a column that only exists
 * once the portal has added it.
 */
export async function saveEmailSettings({ siteUrl, token, settings }) {
  const digest = await provisionSchema(siteUrl, token, {
    lists: [{
      title: SETTINGS_LIST_NAME,
      description: 'IT Service Portal settings shared by everyone in IT',
      columns: SETTINGS_COLUMNS,
    }],
  });
  await provisionLinks(siteUrl, token);

  const value = serialiseEmailSettings(settings);
  const row = await findRow(siteUrl, token, EMAIL_SETTING_KEY);
  const path = row
    ? `${listPath(SETTINGS_LIST_NAME)}/items(${Number(row.Id ?? row.ID)})`
    : `${listPath(SETTINGS_LIST_NAME)}/items`;

  const response = await withRetry(() => spFetch(siteUrl, path, {
    token,
    digest,
    method: 'POST',
    accept: ITEM_ACCEPT,
    body: { Title: EMAIL_SETTING_KEY, SettingValue: value },
    headers: row ? { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' } : undefined,
  }));
  if (!response.ok) {
    throw new Error(`Could not save the email settings (${response.status}): ${await response.text()}`);
  }
  return normaliseEmailSettings(value);
}

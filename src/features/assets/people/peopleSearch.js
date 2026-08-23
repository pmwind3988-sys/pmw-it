import { spFetch } from '../../sharepoint/spClient.js';

/**
 * Finding a real person in the company directory.
 *
 * This uses SharePoint's own people picker rather than Microsoft Graph's
 * `/users`. Graph would also give job title and department, but it needs
 * `User.ReadBasic.All` consented by an admin before anybody can use the feature
 * at all — a poor trade for two fields, and a feature that cannot ship until
 * somebody else acts is a feature that does not ship (§4.5).
 */

const PICKER_PATH = '/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface'
  + '.clientPeoplePickerSearchUser';

/**
 * `PrincipalType: 1` is users only — 15 would also return security groups and
 * distribution lists, and a laptop cannot be handed to a distribution list.
 * `PrincipalSource: 15` searches every source the site knows, which is what
 * finds somebody who has never opened this site before.
 */
function queryBody(query, maximumEntitySuggestions) {
  return {
    queryParams: {
      __metadata: { type: 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
      AllowEmailAddresses: true,
      AllowMultipleEntities: false,
      AllUrlZones: false,
      MaximumEntitySuggestions: maximumEntitySuggestions,
      PrincipalSource: 15,
      PrincipalType: 1,
      QueryString: query,
    },
  };
}

/**
 * The picker answers with a JSON STRING inside its JSON, which is unusual
 * enough to be worth naming: `d.ClientPeoplePickerSearchUser` has to be parsed
 * a second time.
 */
function parseResults(payload) {
  const raw = payload?.d?.ClientPeoplePickerSearchUser;
  if (!raw) return [];

  try {
    const parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

/**
 * Email lives in different places depending on how the account reached the
 * directory: `EntityData.Email` for a directory user, `Description` for one
 * matched by address, and the login name itself for a plain member claim.
 */
export function normalisePerson(entry) {
  const login = entry?.Key ?? '';
  const fromClaim = login.includes('|') ? login.split('|').pop() : login;

  const email = entry?.EntityData?.Email
    || (String(entry?.Description ?? '').includes('@') ? entry.Description : '')
    || (fromClaim.includes('@') ? fromClaim : '');

  return {
    name: entry?.DisplayText || email || login,
    email: String(email).trim().toLowerCase(),
    login,
    title: entry?.EntityData?.Title ?? '',
    department: entry?.EntityData?.Department ?? '',
  };
}

/**
 * Two letters find half the company, so the caller debounces and the search
 * refuses anything shorter than three characters outright.
 */
export const MIN_QUERY = 3;

export async function searchPeople(siteUrl, token, query, { limit = 10 } = {}) {
  const term = String(query ?? '').trim();
  if (term.length < MIN_QUERY) return [];

  const response = await spFetch(siteUrl, PICKER_PATH, {
    token,
    method: 'POST',
    body: queryBody(term, limit),
  });

  if (!response.ok) {
    throw new Error(`The directory search failed (${response.status})`);
  }

  return parseResults(await response.json())
    .map(normalisePerson)
    // Without an email there is nothing to key a person's items on, so an
    // entry that cannot be identified is dropped rather than offered.
    .filter((person) => person.email);
}

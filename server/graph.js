import { ConflictError } from './errors.js';
import { LINKS_LIST_NAME } from '../src/features/forms/links/linkSchema.js';
import {
  CHECKLIST_LIST_NAME, SIGNATURE_LIBRARY_NAME,
} from '../src/features/forms/sharepoint/checklistSchema.js';

/**
 * SharePoint, reached through Microsoft Graph under the portal's OWN identity
 * rather than a signed-in person's.
 *
 * Graph and not SharePoint's REST API, because SharePoint REST refuses an
 * app-only token obtained with a client secret ("Unsupported app only token")
 * and would need a certificate. Graph takes the secret. The app is granted
 * `Sites.Selected` and write on the IThelpdesk site only — see
 * `docs/checklist-links-setup.md`.
 *
 * Exposes exactly the five calls the link handlers need, so the fake in
 * `fakeGraph.js` can stand in for it method for method.
 */

const GRAPH = 'https://graph.microsoft.com/v1.0';

export function createGraph({
  tenantId, clientId, clientSecret, siteUrl, fetchImpl = globalThis.fetch,
}) {
  let token = null;
  let tokenExpires = 0;
  const cache = new Map();

  async function accessToken() {
    if (token && Date.now() < tokenExpires) return token;

    const response = await fetchImpl(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        grant_type: 'client_credentials',
        client_id: clientId,
        client_secret: clientSecret,
        scope: 'https://graph.microsoft.com/.default',
      }),
    });
    if (!response.ok) {
      throw new Error(`Graph sign-in refused (${response.status}): ${await response.text()}`);
    }
    const data = await response.json();
    token = data.access_token;
    // A minute early, so a token never expires halfway through a submission.
    tokenExpires = Date.now() + (Number(data.expires_in) - 60) * 1000;
    return token;
  }

  async function call(path, { method = 'GET', body, headers = {}, raw = false } = {}) {
    const response = await fetchImpl(path.startsWith('http') ? path : `${GRAPH}${path}`, {
      method,
      headers: {
        Authorization: `Bearer ${await accessToken()}`,
        ...(body !== undefined && !raw ? { 'Content-Type': 'application/json' } : null),
        ...headers,
      },
      body: body === undefined ? undefined : (raw ? body : JSON.stringify(body)),
    });

    if (response.status === 412) throw new ConflictError();
    if (!response.ok) {
      throw new Error(`Graph ${method} ${path} failed (${response.status}): ${await response.text()}`);
    }
    return response;
  }

  const once = (key, load) => {
    if (!cache.has(key)) {
      cache.set(key, load().catch((error) => {
        cache.delete(key);
        throw error;
      }));
    }
    return cache.get(key);
  };

  const siteId = () => once('site', async () => {
    const { host, pathname } = new URL(siteUrl);
    const response = await call(`/sites/${host}:${pathname.replace(/\/$/, '')}?$select=id`);
    return (await response.json()).id;
  });

  // Looked up by display name once per instance: Graph addresses lists by id.
  const listId = (title) => once(`list:${title}`, async () => {
    const filter = encodeURIComponent(`displayName eq '${title.replace(/'/g, "''")}'`);
    const response = await call(`/sites/${await siteId()}/lists?$filter=${filter}&$select=id`);
    const found = (await response.json()).value?.[0];
    if (!found) throw new Error(`The list "${title}" does not exist yet — create a link from the portal first`);
    return found.id;
  });

  const signatureDrive = () => once('drive', async () => {
    const response = await call(`/sites/${await siteId()}/lists/${await listId(SIGNATURE_LIBRARY_NAME)}/drive?$select=id`);
    return (await response.json()).id;
  });

  const itemsPath = async (title) => `/sites/${await siteId()}/lists/${await listId(title)}/items`;
  const filePath = async (fileName) => `/drives/${await signatureDrive()}/root:/${encodeURIComponent(fileName)}:`;

  return {
    async findLink(code) {
      // `code` has already passed `isLinkCode`, so it holds no quote to escape.
      const filter = encodeURIComponent(`fields/Title eq '${code}'`);
      const response = await call(`${await itemsPath(LINKS_LIST_NAME)}?$expand=fields&$filter=${filter}&$top=1`, {
        // Title is not indexed on a fresh list; small lists answer regardless.
        headers: { Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' },
      });
      const item = (await response.json()).value?.[0];
      return item ? { id: item.id, eTag: item.eTag, fields: item.fields } : null;
    },

    async updateLink(id, fields, { eTag } = {}) {
      await call(`${await itemsPath(LINKS_LIST_NAME)}/${id}/fields`, {
        method: 'PATCH',
        body: fields,
        headers: eTag ? { 'If-Match': eTag } : {},
      });
    },

    async createChecklist(fields) {
      const response = await call(await itemsPath(CHECKLIST_LIST_NAME), {
        method: 'POST',
        body: { fields },
      });
      return { id: (await response.json()).id };
    },

    async uploadSignature(fileName, bytes) {
      const response = await call(`${await filePath(fileName)}/content`, {
        method: 'PUT',
        body: bytes,
        raw: true,
        headers: { 'Content-Type': 'image/png' },
      });
      const data = await response.json();
      // Stored server-relative, like a signature saved from inside the portal.
      return { serverRelativeUrl: decodeURIComponent(new URL(data.webUrl).pathname) };
    },

    async readSignature(fileName) {
      try {
        const response = await call(`${await filePath(fileName)}/content`);
        return new Uint8Array(await response.arrayBuffer());
      } catch {
        return null;
      }
    },
  };
}

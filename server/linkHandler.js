import { createGraph } from './graph.js';
import { createLinkApi } from './checklistLinkApi.js';

/**
 * HTTP around `createLinkApi`, written against Node's plain request and
 * response so the same function serves Vercel (`api/c/[code].js`) and the Vite
 * dev server (`vite.config.js`).
 */

const DEFAULT_SITE = 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

const HEADERS = {
  'Content-Type': 'application/json; charset=utf-8',
  // The answer is one person's checklist: never cached, never indexed, and the
  // link it came from never passed on to whatever the page links to.
  'Cache-Control': 'no-store',
  'X-Robots-Tag': 'noindex, nofollow',
  'Referrer-Policy': 'no-referrer',
  'X-Content-Type-Options': 'nosniff',
};

export function configFromEnv(env = process.env) {
  const config = {
    tenantId: env.CHECKLIST_TENANT_ID || env.VITE_TENANT_ID,
    clientId: env.CHECKLIST_APP_CLIENT_ID,
    clientSecret: env.CHECKLIST_APP_CLIENT_SECRET,
    siteUrl: env.VITE_SHAREPOINT_SITE_URL || DEFAULT_SITE,
  };
  const missing = [
    ['CHECKLIST_TENANT_ID', config.tenantId],
    ['CHECKLIST_APP_CLIENT_ID', config.clientId],
    ['CHECKLIST_APP_CLIENT_SECRET', config.clientSecret],
  ].filter(([, value]) => !value).map(([name]) => name);
  return { config, missing };
}

let shared = null;

function defaultApi() {
  const { config, missing } = configFromEnv();
  if (missing.length) {
    console.error(`[checklist-link] not configured — missing ${missing.join(', ')}`);
    return null;
  }
  // Kept between requests on a warm instance, so the token and the list ids
  // are looked up once rather than on every visit.
  shared ??= createLinkApi({ graph: createGraph(config) });
  return shared;
}

function send(res, status, body) {
  res.statusCode = status;
  for (const [name, value] of Object.entries(HEADERS)) res.setHeader(name, value);
  res.end(JSON.stringify(body));
}

async function readBody(req) {
  if (req.body !== undefined) {
    return typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
  }
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > 2_000_000) throw new Error('too large');
    chunks.push(chunk);
  }
  return JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}');
}

export async function handleLinkRequest(req, res, code, { api = defaultApi() } = {}) {
  if (!api) {
    send(res, 503, { error: 'This form cannot be opened right now. Please let IT know.' });
    return;
  }

  try {
    if (req.method === 'GET') {
      const { status, body } = await api.get(code);
      send(res, status, body);
      return;
    }

    if (req.method === 'POST') {
      let body;
      try {
        body = await readBody(req);
      } catch {
        send(res, 400, { error: 'The form could not be read. Please try again.' });
        return;
      }
      const { status, body: answer } = await api.submit(code, body);
      send(res, status, answer);
      return;
    }

    res.setHeader('Allow', 'GET, POST');
    send(res, 405, { error: 'Method not allowed' });
  } catch (error) {
    console.error('[checklist-link] request failed', error);
    send(res, 502, { error: 'This form cannot be reached right now. Please try again in a moment.' });
  }
}

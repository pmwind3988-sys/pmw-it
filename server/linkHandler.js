import { createGraph } from './graph.js';
import { createLinkApi } from './checklistLinkApi.js';
import { createChecklistMailer } from './checklistMailer.js';

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
    // The mailbox the signing emails are sent from. Optional: without it,
    // links work exactly as before and nothing is emailed.
    mailFrom: env.CHECKLIST_MAIL_FROM || '',
    // The PORTAL's app registration, whose sign-in tokens may ask for a
    // resend (`mailHandler.js`). Not the server's own.
    portalClientId: env.VITE_CLIENT_ID || '',
  };
  const missing = [
    ['CHECKLIST_TENANT_ID', config.tenantId],
    ['CHECKLIST_APP_CLIENT_ID', config.clientId],
    ['CHECKLIST_APP_CLIENT_SECRET', config.clientSecret],
  ].filter(([, value]) => !value).map(([name]) => name);
  return { config, missing };
}

let shared = null;

/**
 * Graph, the mailer and the link API for these settings, or null with the
 * missing ones logged. Kept between requests on a warm instance, so the token
 * and the list ids are looked up once rather than on every visit.
 */
export function servicesFromEnv(env = process.env) {
  if (shared) return shared;
  const { config, missing } = configFromEnv(env);
  if (missing.length) {
    console.error(`[checklist-link] not configured — missing ${missing.join(', ')}`);
    return null;
  }
  const graph = createGraph(config);
  const mailer = createChecklistMailer({ graph, from: config.mailFrom });
  shared = {
    config, graph, mailer, api: createLinkApi({ graph, mailer }),
  };
  return shared;
}

export const apiFromEnv = (env = process.env) => servicesFromEnv(env)?.api ?? null;

/**
 * The portal's address, for the buttons in an email. Configured first — the
 * Host header of an anonymous request is not something to put in a link IT
 * will click — and the request's own host only when nothing is set (dev).
 */
export function portalOrigin(req, env = process.env) {
  const configured = env.PORTAL_URL || env.VITE_REDIRECT_URI;
  if (configured) {
    try {
      return new URL(configured).origin;
    } catch {
      // Not a URL; fall through to the request.
    }
  }
  const host = req?.headers?.['x-forwarded-host'] || req?.headers?.host || 'localhost:5173';
  const proto = req?.headers?.['x-forwarded-proto'] || (/^localhost|^127./.test(host) ? 'http' : 'https');
  return `${String(proto).split(',')[0]}://${String(host).split(',')[0]}`;
}

export function send(res, status, body) {
  res.statusCode = status;
  for (const [name, value] of Object.entries(HEADERS)) res.setHeader(name, value);
  res.end(JSON.stringify(body));
}

export async function readBody(req) {
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

export async function handleLinkRequest(req, res, code, { api = apiFromEnv(), env = process.env } = {}) {
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
      const { status, body: answer } = await api.submit(code, body, { origin: portalOrigin(req, env) });
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

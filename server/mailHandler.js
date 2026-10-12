import { servicesFromEnv, portalOrigin, send, readBody } from './linkHandler.js';

/**
 * `/api/checklist-email` — somebody in IT asking for a signed checklist to be
 * emailed again, or for a test. Unlike `/api/c/<code>`, this one needs a
 * sign-in: the portal sends the person's Microsoft Graph token, and the
 * server believes it only after Graph itself has accepted it (`/me`) AND the
 * token says it was issued, in this tenant, to the PORTAL's app. A Graph
 * token from any other app or tenant is refused.
 *
 * The request names a link and nothing else. Who receives the email is read
 * from the saved setting on the server, so this cannot be pointed at an
 * address of the caller's choosing.
 */

const GRAPH_AUDIENCES = new Set(['00000003-0000-0000-c000-000000000000', 'https://graph.microsoft.com']);

function claimsOf(token) {
  try {
    const [, payload] = String(token).split('.');
    return JSON.parse(Buffer.from(payload.replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8'));
  } catch {
    return null;
  }
}

/**
 * The caller's name and email, or null. The claims are read before Graph is
 * asked, but only trusted after: Graph rejects a forged or expired token, and
 * the tenant and app checks stop a genuine token minted for somebody else's
 * app from counting here.
 */
export function createCallerCheck({ tenantId, portalClientId, fetchImpl = globalThis.fetch, now = Date.now }) {
  return async function verifyCaller(token) {
    if (!token || !tenantId || !portalClientId) return null;
    const claims = claimsOf(token);
    if (!claims) return null;
    if (claims.tid !== tenantId) return null;
    if ((claims.appid ?? claims.azp) !== portalClientId) return null;
    if (!GRAPH_AUDIENCES.has(claims.aud)) return null;
    if (Number(claims.exp) * 1000 <= now()) return null;

    let response;
    try {
      response = await fetchImpl('https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName', {
        headers: { Authorization: `Bearer ${token}` },
      });
    } catch {
      return null;
    }
    if (!response.ok) return null;
    const me = await response.json();
    return {
      name: me.displayName || me.userPrincipalName || 'IT',
      email: me.mail || me.userPrincipalName || '',
    };
  };
}

const bearer = (req) => {
  const header = req.headers?.authorization ?? req.headers?.Authorization ?? '';
  const match = /^Bearer\s+(.+)$/i.exec(String(header));
  return match ? match[1].trim() : '';
};

export async function handleMailRequest(req, res, { services, verifyCaller, env = process.env } = {}) {
  const ready = services ?? servicesFromEnv(env);
  if (!ready) {
    send(res, 503, { error: 'The portal server is not set up yet. See docs/checklist-links-setup.md.' });
    return;
  }
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    send(res, 405, { error: 'Method not allowed' });
    return;
  }

  const check = verifyCaller ?? createCallerCheck({
    tenantId: ready.config.tenantId, portalClientId: ready.config.portalClientId,
  });
  const caller = await check(bearer(req));
  if (!caller) {
    send(res, 401, { error: 'Your sign-in could not be confirmed. Reload the page and try again.' });
    return;
  }

  let body;
  try {
    body = await readBody(req);
  } catch {
    send(res, 400, { error: 'The request could not be read.' });
    return;
  }

  try {
    const origin = portalOrigin(req, env);
    let result;
    if (body?.action === 'resend') {
      const id = Number(body.id);
      if (!Number.isInteger(id) || id <= 0) {
        send(res, 400, { error: 'Which checklist? The request named none.' });
        return;
      }
      result = await ready.mailer.resend({ id, by: caller.name, origin });
    } else if (body?.action === 'test') {
      result = await ready.mailer.test({ by: caller.name, origin });
    } else {
      send(res, 400, { error: 'Unknown request.' });
      return;
    }
    if (result.ok) send(res, 200, { ok: true, to: result.to });
    else send(res, result.status ?? 502, { ok: false, error: result.error });
  } catch (error) {
    console.error('[checklist-email] request failed', error);
    send(res, 502, { error: 'The email could not be sent just now. Please try again in a moment.' });
  }
}

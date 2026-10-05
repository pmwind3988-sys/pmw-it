import { handleLinkRequest } from '../../server/linkHandler.js';

/**
 * `/api/c/:code` — the only part of the portal that answers without a
 * sign-in. Everything it does is in `server/`; see
 * `docs/checklist-links-setup.md` for the settings it needs.
 */
export default function handler(req, res) {
  return handleLinkRequest(req, res, String(req.query?.code ?? ''));
}

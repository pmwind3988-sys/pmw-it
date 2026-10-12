import { handleMailRequest } from '../server/mailHandler.js';

/**
 * `/api/checklist-email` — IT resending a signed checklist's email, or sending
 * a test. Needs the portal's sign-in; everything it does is in
 * `server/mailHandler.js`.
 */
export default function handler(req, res) {
  return handleMailRequest(req, res);
}

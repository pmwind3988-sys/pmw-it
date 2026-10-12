/**
 * Asking the portal's own server to email a signed checklist again, or to
 * send a test. The server, not the browser, sends: the mail goes out from the
 * portal's mailbox, to the group saved in Email settings, whoever presses the
 * button. The browser only says WHICH link, and proves who is asking with its
 * Graph sign-in token.
 *
 * Answers `{ ok: true, to }` or `{ ok: false, error }` with a sentence fit to
 * show; it does not throw for an answer the server gave.
 */
export async function requestChecklistEmail({ graphToken, action, id }) {
  let response;
  try {
    response = await fetch('/api/checklist-email', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${graphToken}`,
      },
      body: JSON.stringify({ action, id }),
    });
  } catch {
    return { ok: false, error: 'The portal could not be reached. Check the connection and try again.' };
  }

  let body = {};
  try {
    body = await response.json();
  } catch {
    // An HTML error page or nothing at all: the status says enough.
  }
  if (response.ok && body.ok) return { ok: true, to: body.to ?? [] };
  return { ok: false, error: body.error || `The email was not sent (${response.status}).` };
}

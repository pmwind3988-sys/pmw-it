import { provisionLists } from './provisionLists.js';

// The columns a write names only exist once an import has provisioned them,
// and the map -> machine -> edit path can be the very first thing anyone does.
// One run per page load; a failed run is forgotten so the next press retries.
let provisioning = null;

// Resolves to the digest provisioning obtained only for the call that ran it;
// later calls get null and fetch their own, because a digest expires (~30 min)
// and a page left open must not keep sending a dead one.
export async function ensureProvisioned(siteUrl, token) {
  if (provisioning) {
    await provisioning;
    return null;
  }
  const run = provisionLists(siteUrl, token);
  provisioning = run.catch((failure) => {
    provisioning = null;
    throw failure;
  });
  return (await provisioning, await run);
}

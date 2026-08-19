/**
 * Wraps a promise with a timeout.
 * @param {Promise} promise - The promise to wrap
 * @param {number} ms - Timeout in milliseconds (default 30s)
 * @param {string} errorMessage - Custom error message
 * @returns {Promise} - Resolves with original promise or rejects on timeout
 */
export function withTimeout(promise, ms = 30000, errorMessage = 'Request timed out') {
  return Promise.race([
    promise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(errorMessage)), ms)
    )
  ]);
}

/**
 * Wraps MSAL token acquisition with timeout protection.
 * @param {object} msalInstance - The MSAL PublicClientApplication instance
 * @param {object} request - The token request parameters
 * @param {number} timeoutMs - Timeout in milliseconds (default 30s)
 * @returns {Promise} - Token response or timeout error
 */
export async function acquireTokenWithTimeout(msalInstance, request, timeoutMs = 30000) {
  return withTimeout(
    msalInstance.acquireTokenSilent(request),
    timeoutMs,
    'Token acquisition timed out'
  );
}
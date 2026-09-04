'use strict';

class HttpStatusError extends Error {
  constructor(label, status) {
    super(`${label} failed with HTTP ${status}`);
    this.name = 'HttpStatusError';
    this.status = status;
  }
}

function isObject(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

async function discardResponse(response) {
  try {
    await response.body?.cancel();
  } catch {
    // The body may already be consumed or the transport may already be closed.
  }
}

/**
 * Fetch with a token resolved immediately before every network attempt. A 403
 * means the guest token is normally stale, so wait for the realtime provider
 * to publish a different token and retry once. A 401 is not retried because it
 * normally denotes an account-gated endpoint.
 */
async function fetchWithLiveToken(client, tokens, url, init = {}, options = {}) {
  const {
    baseHeaders = {},
    authRetries = 1,
    authRefreshTimeoutMs = 15000,
    clientOptions = {},
  } = options;

  let tokenUsed = null;
  const initFactory = async (context) => {
    const resolvedInit = typeof init === 'function' ? await init(context) : init;
    const token = tokens.token || (await tokens.ready(authRefreshTimeoutMs));
    tokenUsed = token;
    return {
      ...resolvedInit,
      headers: {
        ...baseHeaders,
        ...(resolvedInit.headers || {}),
        u0: token,
      },
    };
  };

  let response = await client.fetch(url, initFactory, clientOptions);
  for (let refresh = 0; response.status === 403 && refresh < authRetries; refresh++) {
    const staleToken = tokenUsed;
    await discardResponse(response);
    await tokens.waitForChange(staleToken, authRefreshTimeoutMs);
    response = await client.fetch(url, initFactory, clientOptions);
  }
  return response;
}

async function readJsonResponse(response, label) {
  const text = await response.text();
  if (!response.ok) throw new HttpStatusError(label, response.status);
  if (!text.trim()) throw new Error(`${label} returned an empty response`);

  let value;
  try {
    value = JSON.parse(text);
  } catch (error) {
    throw new Error(`${label} returned invalid JSON: ${error.message}`);
  }
  if (value === null || (!isObject(value) && !Array.isArray(value))) {
    throw new Error(`${label} returned an unexpected JSON value`);
  }
  if (isObject(value) && value.status && String(value.status).toLowerCase() !== 'success') {
    throw new Error(`${label} returned API status ${value.status}`);
  }
  return value;
}

function requireItems(payload, label) {
  const items = Array.isArray(payload) ? payload : payload?.items;
  if (!Array.isArray(items) || items.length === 0) {
    throw new Error(`${label} returned no items`);
  }
  return items;
}

function isFatalRequestError(error) {
  return /Circuit breaker open|waiting for (?:a )?(?:live|refreshed) token|Token provider stopped/i.test(
    error?.message || ''
  );
}

module.exports = {
  HttpStatusError,
  discardResponse,
  fetchWithLiveToken,
  readJsonResponse,
  requireItems,
  isFatalRequestError,
};

// -----------------------------------------------------------------------------
// Signing in to Microsoft: the OAuth 2.0 device authorization grant of the
// Microsoft identity platform, with the user's own public application.
//
// 1. POST /{tenant}/oauth2/v2.0/devicecode → a short code to type at
//    https://login.microsoft.com/device (the page Microsoft names in
//    `verification_uri`), valid 15 minutes;
// 2. POST /{tenant}/oauth2/v2.0/token every `interval` seconds until the user
//    signed in (`authorization_pending` meanwhile);
// 3. the refresh token (scope `offline_access`) then renews the one-hour
//    access token; Microsoft may hand out a new refresh token each time,
//    which replaces the old one.
//
// No client secret: the application is a public client ("Allow public client
// flows" in its registration). The tokens never reach a log line.
// -----------------------------------------------------------------------------

export const LOGIN_BASE = 'https://login.microsoftonline.com';
export const DEFAULT_VERIFICATION_URI = 'https://microsoft.com/devicelogin';
export const SCOPES = ['offline_access', 'User.Read', 'Calendars.ReadWrite', 'Tasks.ReadWrite'];
const DEVICE_GRANT = 'urn:ietf:params:oauth:grant-type:device_code';
// Renew the access token this long before it expires.
const EARLY_RENEWAL_MS = 5 * 60_000;

// OAuth errors after which only a new sign-in helps.
const SIGN_IN_AGAIN = new Set([
  'invalid_grant',
  'interaction_required',
  'invalid_client',
  'unauthorized_client',
  'consent_required',
  'not_signed_in',
]);

export class AuthError extends Error {
  /**
   * @param {string} code OAuth error code (`invalid_grant`…) or ours
   * @param {{ status?: number, aadsts?: number }} [details]
   */
  constructor(code, { status, aadsts } = {}) {
    super(aadsts ? `${code} (AADSTS${aadsts})` : code);
    this.name = 'AuthError';
    this.code = code;
    this.status = status;
    this.aadsts = aadsts;
  }

  /** True when retrying later cannot help: the user has to sign in again. */
  get needsSignIn() {
    return SIGN_IN_AGAIN.has(this.code);
  }
}

/** The AuthError of a refused response (never its description: it carries ids). */
function errorFrom({ status, body }) {
  const aadsts = Array.isArray(body?.error_codes) ? Number(body.error_codes[0]) : undefined;
  const code = typeof body?.error === 'string' ? body.error : `http_${status}`;
  return new AuthError(code, { status, aadsts: Number.isFinite(aadsts) ? aadsts : undefined });
}

/** The sign-in page Microsoft names, if it is one of its own https pages. */
function safeVerificationUri(value) {
  try {
    const url = new URL(String(value));
    const host = url.hostname.toLowerCase();
    const microsoft = ['microsoft.com', 'microsoftonline.com', 'live.com'].some(
      (domain) => host === domain || host.endsWith(`.${domain}`),
    );
    return url.protocol === 'https:' && microsoft ? url.href : DEFAULT_VERIFICATION_URI;
  } catch {
    return DEFAULT_VERIFICATION_URI;
  }
}

/**
 * @param {object} options
 * @param {ReturnType<import('./store.js').createStore>} options.store
 * @param {typeof fetch} [options.fetch]
 * @param {() => number} [options.now]
 * @param {number} [options.timeoutMs]
 */
export function createAuth({
  store,
  fetch = globalThis.fetch,
  now = Date.now,
  timeoutMs = 20_000,
}) {
  /** @type {{ token: string, expiresAt: number } | null} */
  let access = null;
  let refreshing = null;

  async function post(tenant, endpoint, params) {
    let response;
    try {
      response = await fetch(
        `${LOGIN_BASE}/${encodeURIComponent(tenant)}/oauth2/v2.0/${endpoint}`,
        {
          method: 'POST',
          headers: {
            'content-type': 'application/x-www-form-urlencoded',
            accept: 'application/json',
          },
          body: new URLSearchParams(params).toString(),
          signal: AbortSignal.timeout(timeoutMs),
        },
      );
    } catch (err) {
      // Network error or timeout: its message may hold the address, keep the kind.
      throw new AuthError(err?.name === 'TimeoutError' ? 'timeout' : 'network_error');
    }
    const body = await response.json().catch(() => ({}));
    return { status: response.status, body: body ?? {} };
  }

  /** The stored session, if it belongs to the configured application. */
  function session(config) {
    const tokens = store.getTokens();
    if (!tokens || tokens.client_id !== config.client_id || tokens.tenant !== config.tenant) {
      return null;
    }
    return tokens;
  }

  async function refresh(config, tokens) {
    const answer = await post(config.tenant, 'token', {
      grant_type: 'refresh_token',
      client_id: config.client_id,
      refresh_token: tokens.refresh_token,
      scope: SCOPES.join(' '),
    });
    if (answer.status !== 200 || typeof answer.body.access_token !== 'string') {
      const error = errorFrom(answer);
      if (error.needsSignIn) {
        // Revoked, expired (90 days unused) or a changed password: useless now.
        access = null;
        await store.setTokens(null);
      }
      throw error;
    }
    access = {
      token: answer.body.access_token,
      expiresAt: now() + Number(answer.body.expires_in ?? 3600) * 1000,
    };
    if (
      typeof answer.body.refresh_token === 'string' &&
      answer.body.refresh_token !== tokens.refresh_token
    ) {
      await store.setTokens({ ...tokens, refresh_token: answer.body.refresh_token });
    }
    return access.token;
  }

  return {
    /** True when a refresh token is stored for the configured application. */
    signedIn(config) {
      return session(config) !== null;
    },

    /** The signed-in account (`{ id, name }`), or null. */
    account(config) {
      return session(config)?.account ?? null;
    },

    /**
     * Step 1: ask for a device code.
     * @returns {Promise<{ deviceCode, userCode, verificationUri, expiresAt, intervalMs }>}
     */
    async requestDeviceCode(config) {
      if (!config.client_id) {
        throw new AuthError('missing_client_id');
      }
      const answer = await post(config.tenant, 'devicecode', {
        client_id: config.client_id,
        scope: SCOPES.join(' '),
      });
      if (answer.status !== 200 || typeof answer.body.device_code !== 'string') {
        throw errorFrom(answer);
      }
      return {
        deviceCode: answer.body.device_code,
        userCode: String(answer.body.user_code ?? ''),
        verificationUri: safeVerificationUri(answer.body.verification_uri),
        expiresAt: now() + Number(answer.body.expires_in ?? 900) * 1000,
        intervalMs: Math.max(1, Number(answer.body.interval ?? 5)) * 1000,
      };
    },

    /**
     * Step 2, one attempt: `pending`, `slow_down`, or the tokens once the user
     * signed in. Throws on a refusal, an expired code or a declined consent.
     * @returns {Promise<{ status: 'pending' | 'slow_down' } | { status: 'ok', tokens: object }>}
     */
    async pollDeviceCode(config, deviceCode) {
      const answer = await post(config.tenant, 'token', {
        grant_type: DEVICE_GRANT,
        client_id: config.client_id,
        device_code: deviceCode,
      });
      if (answer.status === 200 && typeof answer.body.access_token === 'string') {
        return { status: 'ok', tokens: answer.body };
      }
      if (answer.body?.error === 'authorization_pending') {
        return { status: 'pending' };
      }
      if (answer.body?.error === 'slow_down') {
        return { status: 'slow_down' };
      }
      throw errorFrom(answer);
    },

    /**
     * Step 3: keep the tokens of a completed sign-in (the access token in
     * memory only), with the account they belong to.
     */
    async adopt(config, tokens, account) {
      if (typeof tokens.refresh_token !== 'string') {
        // Without `offline_access`, the session would end within the hour.
        throw new AuthError('no_refresh_token');
      }
      access = {
        token: tokens.access_token,
        expiresAt: now() + Number(tokens.expires_in ?? 3600) * 1000,
      };
      await store.setTokens({
        client_id: config.client_id,
        tenant: config.tenant,
        refresh_token: tokens.refresh_token,
        account,
      });
    },

    /** The account, once known (the sign-in reads it with the first token). */
    async setAccount(config, account) {
      const tokens = session(config);
      if (tokens) {
        await store.setTokens({ ...tokens, account });
      }
    },

    /**
     * A valid access token, renewed when it expires within 5 minutes (or
     * when `force`: Graph refused the current one).
     */
    async accessToken(config, { force = false } = {}) {
      const tokens = session(config);
      if (!tokens) {
        access = null;
        throw new AuthError('not_signed_in');
      }
      if (!force && access && access.expiresAt - now() > EARLY_RENEWAL_MS) {
        return access.token;
      }
      if (!refreshing) {
        refreshing = refresh(config, tokens).finally(() => {
          refreshing = null;
        });
      }
      return refreshing;
    },

    /** Forget the session (the refresh token is deleted from /data). */
    async signOut() {
      access = null;
      await store.setTokens(null);
    },
  };
}

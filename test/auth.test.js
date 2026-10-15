// -----------------------------------------------------------------------------
// The device code sign-in and the token renewal, against the answers recorded
// from the Microsoft identity platform (test/fixtures/login).
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { AuthError, DEFAULT_VERIFICATION_URI, SCOPES, createAuth } from '../src/auth.js';
import { normalizeConfig } from '../src/config.js';
import { createMemoryStore } from '../src/store.js';
import {
  ACCOUNT,
  CLIENT_ID,
  NOW,
  REFRESH_TOKEN,
  createMicrosoft,
  fixture,
} from './helpers/world.js';

const config = normalizeConfig({ client_id: CLIENT_ID });
const signedInStore = () =>
  createMemoryStore({
    tokens: {
      client_id: CLIENT_ID,
      tenant: 'common',
      refresh_token: REFRESH_TOKEN,
      account: ACCOUNT,
    },
  });

test('a device code: the scopes asked, the page and the code to show', async () => {
  const microsoft = createMicrosoft();
  const auth = createAuth({ store: createMemoryStore(), fetch: microsoft.fetch, now: () => NOW });
  const code = await auth.requestDeviceCode(config);
  assert.deepEqual(code, {
    deviceCode: fixture('login/devicecode.json').device_code,
    userCode: 'ABCDEFGHJ',
    verificationUri: 'https://login.microsoft.com/device',
    expiresAt: NOW + 900_000,
    intervalMs: 5000,
  });
  const [request] = microsoft.requests;
  assert.equal(request.url, 'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode');
  const body = new URLSearchParams(request.body);
  assert.equal(body.get('client_id'), CLIENT_ID);
  assert.deepEqual(body.get('scope').split(' '), SCOPES);
  assert.ok(SCOPES.includes('offline_access'), 'without it, no refresh token');
});

test('an unknown application is refused with its AADSTS code, nothing else', async () => {
  const microsoft = createMicrosoft();
  microsoft.login.deviceCodeError = fixture('login/unknown-client.json');
  const auth = createAuth({ store: createMemoryStore(), fetch: microsoft.fetch });
  await assert.rejects(auth.requestDeviceCode(config), (err) => {
    assert.ok(err instanceof AuthError);
    assert.equal(err.message, 'invalid_request (AADSTS50059)');
    assert.equal(err.needsSignIn, false);
    return true;
  });
  await assert.rejects(auth.requestDeviceCode(normalizeConfig({})), { code: 'missing_client_id' });
});

test('the wait: pending, then the tokens, kept in /data except the access token', async () => {
  const microsoft = createMicrosoft();
  microsoft.login.devicePolls = ['pending'];
  const store = createMemoryStore();
  const auth = createAuth({ store, fetch: microsoft.fetch, now: () => NOW });
  const code = await auth.requestDeviceCode(config);
  assert.deepEqual(await auth.pollDeviceCode(config, code.deviceCode), { status: 'pending' });
  const answer = await auth.pollDeviceCode(config, code.deviceCode);
  assert.equal(answer.status, 'ok');
  await auth.adopt(config, answer.tokens, ACCOUNT);
  assert.deepEqual(store.getTokens(), {
    client_id: CLIENT_ID,
    tenant: 'common',
    refresh_token: microsoft.refreshToken,
    account: ACCOUNT,
  });
  assert.equal(auth.signedIn(config), true);
  assert.equal(await auth.accessToken(config), microsoft.accessToken);
  // The access token is not renewed while it lasts.
  assert.equal(microsoft.requests.length, 3);

  await assert.rejects(auth.pollDeviceCode(config, 'unknown'), {
    code: 'invalid_grant',
    aadsts: 7000014,
  });
});

test('a sign-in without offline_access would end within the hour: refused', async () => {
  const auth = createAuth({ store: createMemoryStore(), fetch: createMicrosoft().fetch });
  await assert.rejects(auth.adopt(config, { access_token: 'x', expires_in: 3599 }, ACCOUNT), {
    code: 'no_refresh_token',
  });
});

test('renewal: before expiry, and the new refresh token replaces the old one', async () => {
  let now = NOW;
  const microsoft = createMicrosoft();
  const store = signedInStore();
  const auth = createAuth({ store, fetch: microsoft.fetch, now: () => now });
  const first = await auth.accessToken(config);
  assert.equal(store.getTokens().refresh_token, microsoft.refreshToken);
  assert.notEqual(store.getTokens().refresh_token, REFRESH_TOKEN);
  // 56 minutes later: still valid for 3 minutes, renewed.
  now += 56 * 60_000;
  const second = await auth.accessToken(config);
  assert.notEqual(second, first);
  // Two callers at once share one renewal.
  const [a, b] = await Promise.all([
    auth.accessToken(config, { force: true }),
    auth.accessToken(config, { force: true }),
  ]);
  assert.equal(a, b);
  assert.equal(microsoft.login.issued, 3);
});

test('a refused refresh token ends the session; a server error does not', async () => {
  const microsoft = createMicrosoft();
  const store = signedInStore();
  const auth = createAuth({ store, fetch: microsoft.fetch });
  microsoft.unreachable = true;
  await assert.rejects(auth.accessToken(config), (err) => {
    assert.equal(err.code, 'network_error');
    assert.equal(err.needsSignIn, false);
    // The message never carries the address nor a token.
    assert.doesNotMatch(err.message, /login|token=/);
    return true;
  });
  assert.equal(auth.signedIn(config), true);

  microsoft.unreachable = false;
  microsoft.login.refreshError = fixture('login/invalid-refresh-token.json');
  await assert.rejects(auth.accessToken(config), (err) => err.needsSignIn);
  assert.equal(store.getTokens(), null);
  assert.equal(auth.signedIn(config), false);
  await assert.rejects(auth.accessToken(config), { code: 'not_signed_in' });
});

test('a session belongs to its application and tenant', async () => {
  const auth = createAuth({ store: signedInStore(), fetch: createMicrosoft().fetch });
  assert.equal(auth.signedIn(config), true);
  assert.equal(
    auth.signedIn(normalizeConfig({ client_id: '11111111-2222-4333-8444-555555555555' })),
    false,
  );
  assert.equal(
    auth.signedIn(normalizeConfig({ client_id: CLIENT_ID, tenant: 'consumers' })),
    false,
  );
  assert.deepEqual(auth.account(config), ACCOUNT);
  await auth.signOut();
  assert.equal(auth.signedIn(config), false);
});

test('only a Microsoft https page is shown as the sign-in page', async () => {
  for (const [uri, shown] of [
    ['https://microsoft.com/devicelogin', 'https://microsoft.com/devicelogin'],
    ['https://login.microsoft.com/device', 'https://login.microsoft.com/device'],
    ['http://login.microsoft.com/device', DEFAULT_VERIFICATION_URI],
    ['https://login.microsoft.com.evil.example/device', DEFAULT_VERIFICATION_URI],
    [undefined, DEFAULT_VERIFICATION_URI],
  ]) {
    const fetch = async () =>
      new Response(JSON.stringify({ ...fixture('login/devicecode.json'), verification_uri: uri }), {
        status: 200,
      });
    const auth = createAuth({ store: createMemoryStore(), fetch });
    assert.equal((await auth.requestDeviceCode(config)).verificationUri, shown, String(uri));
  }
});

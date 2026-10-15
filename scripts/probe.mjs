#!/usr/bin/env node
// -----------------------------------------------------------------------------
// Probe of the real Microsoft endpoints, with the integration's own client
// (src/auth.js, src/graph.js), read-only, printing nothing sensitive.
//
//   node scripts/probe.mjs
//     1. asks a device code with an INVENTED client ID, on the three tenant
//        aliases: Microsoft must refuse it (the AADSTS code is printed);
//     2. calls Graph without a valid token: Graph must answer 401.
//
//   PROBE_CLIENT_ID=<your app id> node scripts/probe.mjs
//     additionally asks a real device code for that application and prints
//     the page and its expiry (not the code), then stops: nothing is signed
//     in, nothing is stored.
//
// Exit code 0 when every answer has the expected shape, 1 otherwise.
// -----------------------------------------------------------------------------

import { randomUUID } from 'node:crypto';
import { AuthError, createAuth } from '../src/auth.js';
import { normalizeConfig } from '../src/config.js';
import { GRAPH_BASE, createGraphClient } from '../src/graph.js';
import { createMemoryStore } from '../src/store.js';

process.env.LOG_LEVEL ??= 'warn';
let failures = 0;

function report(ok, what, detail) {
  if (!ok) {
    failures += 1;
  }
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${what}: ${detail}`);
}

// 1. An invented application, refused by the identity platform.
const invented = randomUUID();
for (const tenant of ['common', 'organizations', 'consumers']) {
  const config = normalizeConfig({ client_id: invented, tenant });
  const auth = createAuth({ store: createMemoryStore() });
  try {
    await auth.requestDeviceCode(config);
    report(false, `devicecode (${tenant}, invented client)`, 'accepted?!');
  } catch (err) {
    const refused = err instanceof AuthError && err.status >= 400 && err.status < 600;
    report(refused, `devicecode (${tenant}, invented client)`, `refused, ${err.message}`);
  }
}

// 2. Graph without a valid token.
const graph = createGraphClient({
  auth: { accessToken: async () => 'invalid-token-for-the-probe' },
  getConfig: () => ({}),
});
try {
  await graph.me();
  report(false, `GET ${GRAPH_BASE}/me without token`, 'answered?!');
} catch (err) {
  report(err.status === 401, `GET ${GRAPH_BASE}/me without token`, err.message);
}

// 3. Optional: a real device code for the user's own application.
if (process.env.PROBE_CLIENT_ID) {
  const config = normalizeConfig({
    client_id: process.env.PROBE_CLIENT_ID,
    tenant: process.env.PROBE_TENANT ?? 'common',
  });
  const auth = createAuth({ store: createMemoryStore() });
  try {
    const code = await auth.requestDeviceCode(config);
    const minutes = Math.round((code.expiresAt - Date.now()) / 60_000);
    report(
      code.userCode.length > 0,
      'devicecode (your application)',
      `page ${code.verificationUri}, code valid ${minutes} min, poll every ${code.intervalMs / 1000} s`,
    );
  } catch (err) {
    report(false, 'devicecode (your application)', err.message);
  }
}

process.exit(failures === 0 ? 0 : 1);

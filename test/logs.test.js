// -----------------------------------------------------------------------------
// What the logs show: nothing at the info level for a normal cycle, one line
// when a problem starts and one when it ends, and never a secret — not even at
// the debug level.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { CLIENT_ID, REFRESH_TOKEN, createWorld, fixture } from './helpers/world.js';

const MINUTE = 60_000;

/** Capture what the SDK logger writes, at `level`, until `stop()`. */
function captureLogs(t, level = 'info') {
  const lines = [];
  const previous = process.env.LOG_LEVEL;
  process.env.LOG_LEVEL = level;
  for (const stream of [process.stdout, process.stderr]) {
    t.mock.method(stream, 'write', (chunk) => {
      lines.push(...String(chunk).split('\n').filter(Boolean));
      return true;
    });
  }
  return {
    lines,
    stop() {
      t.mock.restoreAll();
      process.env.LOG_LEVEL = previous;
    },
  };
}

test('normal cycles write nothing at the info level', async (t) => {
  const world = await createWorld({ config: { trigger_keywords: 'dentiste' } });
  const logs = captureLogs(t);
  try {
    // A whole day: reads, states, scene events, the midnight overdue task.
    await world.run(18 * 60 * MINUTE, MINUTE);
  } finally {
    logs.stop();
    world.app.stop();
  }
  assert.ok(world.gladys.fake.sceneEvents.length > 10, 'the day had events');
  assert.deepEqual(logs.lines, []);
});

test('an outage: one line when it starts, one when it ends', async (t) => {
  const world = await createWorld();
  const logs = captureLogs(t);
  try {
    world.microsoft.unreachable = true;
    await world.run(30 * MINUTE, 5 * MINUTE);
    world.microsoft.unreachable = false;
    await world.run(30 * MINUTE, 5 * MINUTE);
  } finally {
    logs.stop();
    world.app.stop();
  }
  assert.equal(logs.lines.length, 2, logs.lines.join('\n'));
  assert.match(logs.lines[0], /\[WARN\].*Microsoft 365 unreachable \(network_error\)/);
  assert.match(logs.lines[1], /\[INFO\].*Microsoft 365 reachable again/);
});

test('no token, code or client ID in the logs, even at the debug level', async (t) => {
  const logs = captureLogs(t, 'debug');
  const secrets = [];
  try {
    const world = await createWorld({ signedIn: false });
    await world.gladys.fake.action('sign_in', {});
    await world.monitor.signInDone;
    await world.addDevices();
    await world.run(10 * MINUTE, MINUTE);
    // A revoked session and a throttled Graph, logged with their codes only.
    world.microsoft.graphFailures.push({
      status: 429,
      body: fixture('graph/throttled.json'),
      headers: { 'retry-after': '120' },
    });
    await world.advance(5 * MINUTE);
    await world.advance(5 * MINUTE);
    world.microsoft.login.refreshError = fixture('login/invalid-refresh-token.json');
    world.microsoft.accessToken = 'revoked';
    await world.advance(5 * MINUTE);
    secrets.push(
      REFRESH_TOKEN,
      world.microsoft.refreshToken,
      fixture('login/devicecode.json').device_code,
      'eyJ0eXAiOiJKV1Q',
      'M.C5_BAY',
      CLIENT_ID,
    );
    world.app.stop();
  } finally {
    logs.stop();
  }
  const text = logs.lines.join('\n');
  assert.match(text, /\[DEBUG\]/, 'the debug level shows more');
  for (const secret of secrets) {
    assert.ok(!text.includes(secret), `a secret leaked: ${secret.slice(0, 12)}…`);
  }
  assert.match(text, /Microsoft session lost \(invalid_grant \(AADSTS9002313\)\)/);
  assert.match(text, /throttles the integration for 2 min/);
});

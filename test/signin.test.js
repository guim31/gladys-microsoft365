// -----------------------------------------------------------------------------
// The "Sign in", "Check the connection" and "Sign out" actions, as the
// administrator uses them from the Configuration screen.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { CLIENT_ID, createWorld, fixture } from './helpers/world.js';

test('without a client ID: the status and the action say what to do', async () => {
  const world = await createWorld({ signedIn: false, config: { client_id: '' } });
  try {
    assert.equal(world.gladys.fake.connectionStatus.connected, false);
    assert.match(world.gladys.fake.connectionStatus.message.fr, /ID d'application/);
    const ack = await world.gladys.fake.action('sign_in', {});
    assert.equal(ack.success, true);
    assert.match(JSON.stringify(ack.data), /client\) ID/);
    assert.deepEqual(world.microsoft.requests, []);
  } finally {
    world.app.stop();
  }
});

test('sign in: the code to type, the wait, then the devices to add', async () => {
  const world = await createWorld({ signedIn: false });
  try {
    assert.deepEqual(world.gladys.fake.discoveredDevices, []);
    assert.match(world.gladys.fake.connectionStatus.message.en, /Not signed in/);

    const ack = await world.gladys.fake.action('sign_in', {});
    assert.equal(ack.success, true);
    const message = ack.data.message ?? ack.data;
    assert.equal(
      message.fr,
      'Ouvrez https://login.microsoft.com/device et saisissez le code ABCDEFGHJ dans les 15 minutes, puis connectez-vous avec votre compte Microsoft.',
    );
    assert.match(message.de, /Öffne .* gib .* ein und melde dich/);
    assert.match(message.es, /Abra .* introduzca el código ABCDEFGHJ/);
    // While waiting, the Configuration screen shows the code too.
    assert.match(world.gladys.fake.connectionStatus.message.en, /enter the code ABCDEFGHJ/);
    // A second click gives the same code, without asking Microsoft again.
    const again = await world.gladys.fake.action('sign_in', {});
    assert.match(JSON.stringify(again.data), /ABCDEFGHJ/);
    assert.equal(
      world.microsoft.requests.filter((request) => request.url.endsWith('/devicecode')).length,
      1,
    );

    await world.monitor.signInDone;
    assert.deepEqual(world.gladys.fake.connectionStatus, { connected: true });
    assert.deepEqual(world.store.getTokens(), {
      client_id: CLIENT_ID,
      tenant: 'common',
      refresh_token: world.microsoft.refreshToken,
      account: { id: '00000000-0000-0000-0000-000000000001', name: 'Camille Martin' },
    });
    assert.deepEqual(
      world.gladys.fake.discoveredDevices.map((device) => device.name),
      ['Calendar (Outlook)', 'Famille (Outlook)', 'Tasks (To Do)', 'Monthly Tasks (To Do)'],
    );
    // The token endpoint was asked every 5 seconds (twice pending, then ok).
    const polls = world.microsoft.requests.filter(
      (request) => request.url.endsWith('/token') && String(request.body).includes('device_code'),
    );
    assert.equal(polls.length, 3);

    // The check reports the account and what it holds.
    const check = await world.gladys.fake.action('check_connection', {});
    assert.equal(
      (check.data.message ?? check.data).en,
      'Signed in as Camille Martin: 2 calendar(s), 2 To Do list(s). Add them in the “Discover” tab.',
    );
  } finally {
    world.app.stop();
  }
});

test('the code expires, or the user declines: the status says so', async () => {
  for (const [answer, expected] of [
    ['expired_token', /expired/],
    ['authorization_declined', /declined/],
  ]) {
    const world = await createWorld({ signedIn: false });
    try {
      world.microsoft.login.devicePolls = ['pending', answer];
      await world.gladys.fake.action('sign_in', {});
      await world.monitor.signInDone;
      assert.equal(world.gladys.fake.connectionStatus.connected, false);
      assert.match(world.gladys.fake.connectionStatus.message.en, expected);
      assert.equal(world.store.getTokens(), null);
    } finally {
      world.app.stop();
    }
  }
});

test('nobody types the code: the wait stops after 15 minutes', async () => {
  const world = await createWorld({ signedIn: false });
  try {
    world.microsoft.login.devicePolls = Array(1000).fill('pending');
    await world.gladys.fake.action('sign_in', {});
    await world.monitor.signInDone;
    assert.match(world.gladys.fake.connectionStatus.message.en, /code expired/);
    const polls = world.microsoft.requests.filter((request) => request.url.endsWith('/token'));
    assert.equal(polls.length, 180, '15 minutes every 5 seconds');
  } finally {
    world.app.stop();
  }
});

test('Microsoft refuses the application: the AADSTS code is shown', async () => {
  const world = await createWorld({ signedIn: false });
  try {
    world.microsoft.login.deviceCodeError = fixture('login/unknown-client.json');
    const ack = await world.gladys.fake.action('sign_in', {});
    assert.match(
      (ack.data.message ?? ack.data).en,
      /refused the sign-in \(invalid_request \(AADSTS50059\)\)/,
    );
  } finally {
    world.app.stop();
  }
});

test('sign out: tokens deleted, nothing left to discover', async () => {
  const world = await createWorld();
  try {
    const ack = await world.gladys.fake.action('sign_out', {});
    assert.match((ack.data.message ?? ack.data).fr, /Déconnecté/);
    assert.equal(world.store.getTokens(), null);
    assert.deepEqual(world.gladys.fake.discoveredDevices, []);
    assert.equal(world.gladys.fake.connectionStatus.connected, false);
    const check = await world.gladys.fake.action('check_connection', {});
    assert.match((check.data.message ?? check.data).en, /Not signed in/);
  } finally {
    world.app.stop();
  }
});

test('another client ID ends the session it does not own', async () => {
  const world = await createWorld();
  try {
    await world.gladys.fake.configUpdated({
      client_id: '11111111-2222-4333-8444-555555555555',
    });
    assert.equal(world.gladys.fake.connectionStatus.connected, false);
    assert.match(world.gladys.fake.connectionStatus.message.en, /Not signed in/);
  } finally {
    world.app.stop();
  }
});

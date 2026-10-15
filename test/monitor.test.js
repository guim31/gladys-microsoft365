// -----------------------------------------------------------------------------
// The monitor end to end: the SDK's fake Gladys, a fake Microsoft, a clock.
// States, the refresh without the core's poll, the scene triggers, the
// overdue tasks, and what happens when Microsoft says no.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createMemoryStore } from '../src/store.js';
import { NOW, createWorld, fixture, settle } from './helpers/world.js';

const MINUTE = 60_000;
const at = (iso) => Date.parse(iso);

test('a device added to Gladys gets its states at once', async () => {
  const world = await createWorld();
  try {
    assert.equal(world.state('Calendar', 'busy'), 0);
    assert.equal(world.state('Calendar', 'out_of_office'), 0);
    assert.equal(world.state('Calendar', 'current_event'), 'Télétravail');
    assert.equal(world.state('Calendar', 'next_event'), "Point d'équipe · 09:30");
    assert.equal(world.state('Famille', 'current_event'), 'Aucun');
    assert.equal(world.state('Famille', 'next_event'), 'Aucun');
    assert.equal(world.state('Tasks', 'open_tasks'), 4);
    assert.equal(world.state('Tasks', 'overdue_tasks'), 2);
    assert.equal(world.state('Tasks', 'due_today_tasks'), 1);
    assert.equal(world.state('Monthly Tasks', 'open_tasks'), 0);
    assert.deepEqual(world.gladys.fake.connectionStatus, { connected: true });
  } finally {
    world.app.stop();
  }
});

test('only the devices the user added are read', async () => {
  const world = await createWorld({ add: ['Tasks'] });
  try {
    assert.deepEqual(world.microsoft.graphCalls(), [
      'GET /me/calendars',
      'GET /me/todo/lists',
      'GET /me/todo/lists/AAMkADIyAAAAABrJAAA=/tasks',
    ]);
    world.microsoft.requests.length = 0;
    await world.advance(5 * MINUTE);
    assert.deepEqual(world.microsoft.graphCalls(), [
      'GET /me/todo/lists/AAMkADIyAAAAABrJAAA=/tasks',
    ]);
  } finally {
    world.app.stop();
  }
});

test('the states refresh on the container timer, without the core polling', async (t) => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const world = await createWorld({ timers: true });
  try {
    const polls = world.gladys.fake.wsMessages.filter((message) => /poll/.test(message.type));
    assert.equal(polls.length, 0);
    world.microsoft.requests.length = 0;
    const steps = async (count) => {
      for (let step = 0; step < count; step += 1) {
        world.clock.now += 30_000;
        t.mock.timers.tick(30_000);
        await world.monitor.idle();
        await settle();
      }
    };
    // 08:00 → 08:31: Graph read every 5 minutes, 6 times for each of the 4 devices.
    await steps(63);
    assert.equal(world.microsoft.graphCalls().length, 6 * 4);
    assert.equal(world.state('Calendar', 'busy'), 0);
    // 09:31: the meeting started, from what was read (no extra Graph call needed).
    world.clock.now = Date.parse('2026-10-12T07:29:30Z');
    await steps(3);
    assert.equal(world.state('Calendar', 'busy'), 1);
    assert.equal(world.state('Calendar', 'current_event'), "Point d'équipe");
    assert.equal(world.state('Calendar', 'next_event'), 'Dentiste · 14:00');
    // A task added in To Do shows within one period.
    world.microsoft.data.tasks['AAMkADIyAAAAABrJAAA='].push({
      id: 'task-new',
      title: 'Appeler le plombier',
      status: 'notStarted',
    });
    await steps(10);
    assert.equal(world.state('Tasks', 'open_tasks'), 5);
  } finally {
    world.app.stop();
  }
});

test('only changes are published, plus a heartbeat every 6 hours', async (t) => {
  // The SDK dates its heartbeat with Date.now: move it with the clock.
  t.mock.timers.enable({ apis: ['Date'], now: NOW });
  const world = await createWorld({ add: ['Monthly Tasks'] });
  world.run = async (ms, step) => {
    for (let elapsed = 0; elapsed < ms; elapsed += step) {
      world.clock.now += step;
      t.mock.timers.setTime(world.clock.now);
      await world.monitor.tick();
    }
  };
  try {
    const feature = `${world.device('Monthly Tasks').external_id}:open_tasks`;
    const count = () =>
      world.gladys.fake.states.filter((state) => state.device_feature_external_id === feature)
        .length;
    assert.equal(count(), 1);
    await world.run(5 * 3_600_000, 5 * MINUTE);
    assert.equal(count(), 1, 'an unchanged value is not sent again');
    await world.run(1.5 * 3_600_000, 5 * MINUTE);
    assert.equal(count(), 2, 'heartbeat');
  } finally {
    world.app.stop();
  }
});

test('scene events: starts, ends and lead times, with the keyword found', async () => {
  const world = await createWorld({ config: { trigger_keywords: 'dentiste, réunion' } });
  try {
    // Connecting at 08:00 during the remote-work day announces nothing.
    assert.deepEqual(world.gladys.fake.sceneEvents, []);
    const calendar = world.device('Calendar').external_id;

    // 09:15 → 09:31: the 15, 10 and 5 minute lead times, then the start.
    world.clock.now = at('2026-10-12T07:14:30Z');
    await world.monitor.tick();
    await world.run(16 * MINUTE + 30_000);
    const fired = world.gladys.fake.sceneEvents;
    assert.deepEqual(
      fired.map((event) => [event.key, event.data.minutes]),
      [
        ['event_upcoming', 15],
        ['event_upcoming', 10],
        ['event_upcoming', 5],
        ['event_started', undefined],
      ],
    );
    assert.deepEqual(fired.at(-1).data, {
      calendar,
      calendar_name: 'Calendar (Outlook)',
      keyword: null,
      title: "Point d'équipe",
      location: 'Salle Rivoli',
      start: '09:30',
      end: '10:00',
    });

    // 13:59 → 14:46: the dentist starts and ends, with its keyword.
    world.clock.now = at('2026-10-12T11:59:00Z');
    await world.monitor.tick();
    await world.run(47 * MINUTE);
    // (The 13:55 lead time, 4 minutes before the jump, was announced too.)
    const dentist = world.gladys.fake.sceneEvents.filter(
      (event) => event.data.title === 'Dentiste' && event.key !== 'event_upcoming',
    );
    assert.deepEqual(
      dentist.map((event) => [event.key, event.data.keyword]),
      [
        ['event_started', 'dentiste'],
        ['event_ended', 'dentiste'],
      ],
    );
    // Busy and out of office during the appointment, back after.
    assert.equal(world.state('Calendar', 'busy'), 0);
    assert.equal(world.state('Calendar', 'out_of_office'), 0);
    assert.ok(
      world.gladys.fake.states.some(
        (state) =>
          state.device_feature_external_id === `${calendar}:out_of_office` && state.state === 1,
      ),
    );
  } finally {
    world.app.stop();
  }
});

test('an instant missed by more than 5 minutes is not announced late', async () => {
  const world = await createWorld();
  try {
    // The container was paused from 09:00 to 09:40: the 09:30 start is lost.
    world.clock.now = at('2026-10-12T07:00:00Z');
    await world.monitor.tick();
    const before = world.gladys.fake.sceneEvents.length;
    world.clock.now = at('2026-10-12T07:40:00Z');
    await world.monitor.tick();
    const late = world.gladys.fake.sceneEvents.slice(before).map((event) => event.key);
    assert.deepEqual(late, []);
  } finally {
    world.app.stop();
  }
});

test('the all-day event ends at midnight; tasks become overdue, announced once', async () => {
  const store = createMemoryStore({
    tokens: {
      client_id: '0f1e2d3c-4b5a-4968-8776-655443322110',
      tenant: 'common',
      refresh_token: 'M.C5_BAY.0.U.-Cg9refresh-token-invented-for-the-tests!',
      account: { id: '00000000-0000-0000-0000-000000000001', name: 'Camille Martin' },
    },
  });
  const world = await createWorld({ store });
  try {
    // Already overdue at the first read: not news.
    assert.deepEqual(store.state.overdue[Object.keys(store.state.overdue)[0]].sort(), [
      'task-bins',
      'task-plants',
    ]);
    world.clock.now = at('2026-10-12T21:58:00Z');
    await world.monitor.tick();
    await world.run(3 * MINUTE);
    const keys = world.gladys.fake.sceneEvents.map((event) => event.key);
    assert.ok(keys.includes('event_ended'), 'Télétravail ends at midnight');
    const overdue = world.gladys.fake.sceneEvents.filter((event) => event.key === 'task_overdue');
    assert.deepEqual(
      overdue.map((event) => event.data),
      [
        {
          list: world.device('Tasks').external_id,
          list_name: 'Tasks (To Do)',
          title: 'Acheter du pain',
          due: 'lun. 12 oct.',
        },
      ],
    );
    assert.equal(world.state('Tasks', 'overdue_tasks'), 3);
    assert.equal(world.state('Tasks', 'due_today_tasks'), 0);
    await world.run(10 * MINUTE);
    assert.equal(
      world.gladys.fake.sceneEvents.filter((event) => event.key === 'task_overdue').length,
      1,
    );
  } finally {
    world.app.stop();
  }

  // A restart with the same /data does not announce it again.
  const restarted = await createWorld({ store });
  try {
    restarted.clock.now = at('2026-10-12T22:10:00Z');
    await restarted.run(MINUTE);
    assert.deepEqual(
      restarted.gladys.fake.sceneEvents.filter((event) => event.key === 'task_overdue'),
      [],
    );
  } finally {
    restarted.app.stop();
  }
});

test('a revoked session: one warning, the status says to sign in again', async () => {
  const world = await createWorld();
  try {
    world.microsoft.login.refreshError = fixture('login/invalid-refresh-token.json');
    world.microsoft.accessToken = 'revoked';
    await world.advance(5 * MINUTE);
    const status = world.gladys.fake.connectionStatus;
    assert.equal(status.connected, false);
    assert.match(status.message.fr, /expiré ou a été révoquée/);
    assert.equal(world.store.getTokens(), null);
    world.microsoft.requests.length = 0;
    await world.advance(5 * MINUTE);
    assert.deepEqual(world.microsoft.requests, [], 'no more calls without a session');
    // States computed from what was read keep going.
    assert.equal(world.state('Tasks', 'open_tasks'), 4);
  } finally {
    world.app.stop();
  }
});

test('Microsoft unreachable: the status says so, and comes back', async () => {
  const world = await createWorld();
  try {
    world.microsoft.unreachable = true;
    await world.advance(5 * MINUTE);
    assert.equal(world.gladys.fake.connectionStatus.connected, false);
    assert.match(world.gladys.fake.connectionStatus.message.en, /unreachable/);
    world.microsoft.unreachable = false;
    await world.advance(5 * MINUTE);
    assert.deepEqual(world.gladys.fake.connectionStatus, { connected: true });
  } finally {
    world.app.stop();
  }
});

test('Graph throttling: no read until Retry-After is over', async () => {
  const world = await createWorld({ add: ['Tasks'] });
  try {
    world.microsoft.graphFailures.push({
      status: 429,
      body: fixture('graph/throttled.json'),
      headers: { 'retry-after': '900' },
    });
    world.microsoft.requests.length = 0;
    await world.advance(5 * MINUTE);
    assert.equal(world.microsoft.graphCalls().length, 1);
    await world.advance(5 * MINUTE);
    await world.advance(5 * MINUTE);
    assert.equal(world.microsoft.graphCalls().length, 1, 'blocked for 15 minutes');
    await world.advance(5 * MINUTE);
    assert.equal(world.microsoft.graphCalls().length, 2);
    assert.deepEqual(world.gladys.fake.connectionStatus, { connected: true });
  } finally {
    world.app.stop();
  }
});

test('a calendar deleted in Outlook stops being read, the others go on', async () => {
  const world = await createWorld();
  try {
    world.microsoft.data.calendars = world.microsoft.data.calendars.filter(
      (calendar) => calendar.name !== 'Famille',
    );
    await world.advance(5 * MINUTE);
    assert.ok(!world.monitor.cache.has(world.device('Famille').external_id));
    assert.ok(world.monitor.cache.has(world.device('Calendar').external_id));
    assert.deepEqual(world.gladys.fake.connectionStatus, { connected: true });
  } finally {
    world.app.stop();
  }
});

test("the core's poll is a fallback: it reads Graph only when the timer did not", async () => {
  const world = await createWorld();
  try {
    const device = world.gladys.devices.find((entry) => entry.name.startsWith('Tasks'));
    world.microsoft.requests.length = 0;
    world.clock.now += MINUTE;
    assert.deepEqual(await world.gladys.fake.poll(device), { success: true });
    assert.deepEqual(world.microsoft.requests, []);
    world.clock.now += 5 * MINUTE;
    await world.gladys.fake.poll(device);
    assert.equal(world.microsoft.graphCalls().length, 4);
  } finally {
    world.app.stop();
  }
});

test('the widget is nudged when what it shows changes, not otherwise', async () => {
  const world = await createWorld();
  try {
    const nudges = () => world.gladys.fake.widgetRefreshes.length;
    const before = nudges();
    await world.advance(5 * MINUTE);
    assert.equal(nudges(), before);
    world.microsoft.data.tasks['AAMkADIyAAAAABrJAAA='].push({
      id: 'task-new',
      title: 'Appeler le plombier',
      status: 'notStarted',
    });
    await world.advance(5 * MINUTE);
    assert.equal(nudges(), before + 1);
    assert.equal(NOW < world.clock.now, true);
  } finally {
    world.app.stop();
  }
});

// -----------------------------------------------------------------------------
// The Graph client: throttling (Retry-After), a refused access token, paging,
// one request at a time.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createAuth } from '../src/auth.js';
import { normalizeConfig } from '../src/config.js';
import { GraphError, createGraphClient } from '../src/graph.js';
import { createMemoryStore } from '../src/store.js';
import {
  ACCOUNT,
  CLIENT_ID,
  NOW,
  REFRESH_TOKEN,
  createMicrosoft,
  fixture,
} from './helpers/world.js';

function setup(options = {}) {
  const clock = { now: NOW, slept: [] };
  const microsoft = createMicrosoft(options);
  const config = normalizeConfig({ client_id: CLIENT_ID });
  const auth = createAuth({
    store: createMemoryStore({
      tokens: {
        client_id: CLIENT_ID,
        tenant: 'common',
        refresh_token: REFRESH_TOKEN,
        account: ACCOUNT,
      },
    }),
    fetch: microsoft.fetch,
    now: () => clock.now,
  });
  const graph = createGraphClient({
    auth,
    getConfig: () => config,
    fetch: microsoft.fetch,
    now: () => clock.now,
    sleep: async (ms) => {
      clock.slept.push(ms);
      clock.now += ms;
    },
  });
  return { clock, microsoft, graph };
}

const throttled = (seconds) => ({
  status: 429,
  body: fixture('graph/throttled.json'),
  headers: { 'retry-after': String(seconds) },
});

test('a short Retry-After is waited, then the request goes through', async () => {
  const { clock, microsoft, graph } = setup();
  await graph.todoLists(); // signs in
  microsoft.graphFailures.push(throttled(10));
  const lists = await graph.todoLists();
  assert.equal(lists.length, 2);
  assert.deepEqual(clock.slept, [10_000]);
});

test('a long Retry-After blocks every call without asking Graph again', async () => {
  const { clock, microsoft, graph } = setup();
  await graph.todoLists();
  microsoft.graphFailures.push(throttled(600));
  await assert.rejects(graph.todoLists(), (err) => {
    assert.ok(err instanceof GraphError);
    assert.equal(err.code, 'TooManyRequests');
    assert.equal(err.retryAt, NOW + 600_000);
    return true;
  });
  const asked = microsoft.requests.length;
  await assert.rejects(graph.calendars(), { code: 'throttled' });
  assert.equal(microsoft.requests.length, asked, 'no request while blocked');
  assert.equal(graph.blockedUntil, NOW + 600_000);
  clock.now += 600_000;
  assert.equal(graph.blockedUntil, 0);
  assert.equal((await graph.calendars()).length, 2);
});

test('503 without Retry-After: exponential waits, then blocked', async () => {
  const { clock, microsoft, graph } = setup();
  await graph.todoLists();
  microsoft.graphFailures.push(
    { status: 503, body: {} },
    { status: 503, body: {} },
    { status: 503, body: {} },
  );
  await assert.rejects(graph.todoLists(), { status: 503 });
  assert.deepEqual(clock.slept, [2000, 4000]);
});

test('an access token refused before its time is renewed once', async () => {
  const { microsoft, graph } = setup();
  await graph.todoLists();
  microsoft.accessToken = 'revoked-on-the-server-side';
  assert.equal((await graph.todoLists()).length, 2);
  assert.equal(microsoft.login.issued, 2);
});

test('other errors keep their Graph code', async () => {
  const { microsoft, graph } = setup();
  await assert.rejects(graph.openTasks('missing-list'), { status: 404, code: 'ErrorItemNotFound' });
  microsoft.graphFailures.push({
    status: 403,
    body: { error: { code: 'ErrorAccessDenied', message: 'Access is denied.' } },
  });
  await assert.rejects(graph.calendars(), { status: 403, code: 'ErrorAccessDenied' });
  microsoft.unreachable = true;
  await assert.rejects(graph.calendars(), { status: 0, code: 'network_error' });
});

test('every page is read, following only Graph links', async () => {
  const lists = Array.from({ length: 5 }, (_, index) => ({
    id: `list-${index}`,
    displayName: `List ${index}`,
  }));
  const { microsoft, graph } = setup({ lists });
  microsoft.pageSize = 2;
  assert.equal((await graph.todoLists()).length, 5);
  assert.equal(microsoft.graphCalls().length, 3);

  // A next link to another host is not followed (the token would go with it).
  const foreign = setup();
  const original = foreign.microsoft.fetch;
  const graph2 = createGraphClient({
    auth: { accessToken: async () => foreign.microsoft.accessToken ?? 'x' },
    getConfig: () => ({}),
    fetch: async (url, init) => {
      if (String(url).includes('/me/todo/lists')) {
        return new Response(
          JSON.stringify({ value: [{ id: 'a' }], '@odata.nextLink': 'https://evil.example/next' }),
          { status: 200 },
        );
      }
      return original(url, init);
    },
  });
  assert.deepEqual(await graph2.todoLists(), [{ id: 'a' }]);
});

test('one request at a time, in call order', async () => {
  let running = 0;
  let most = 0;
  const order = [];
  const graph = createGraphClient({
    auth: { accessToken: async () => 'token' },
    getConfig: () => ({}),
    fetch: async (url) => {
      running += 1;
      most = Math.max(most, running);
      await new Promise((resolve) => setTimeout(resolve, 5));
      order.push(new URL(url).pathname);
      running -= 1;
      return new Response(JSON.stringify({ value: [] }), { status: 200 });
    },
  });
  await Promise.all([graph.calendars(), graph.todoLists(), graph.openTasks('x')]);
  assert.equal(most, 1);
  assert.deepEqual(order, [
    '/v1.0/me/calendars',
    '/v1.0/me/todo/lists',
    '/v1.0/me/todo/lists/x/tasks',
  ]);
});

test('calendarView asks UTC times, the needed fields and the window', async () => {
  const { microsoft, graph } = setup();
  const events = await graph.calendarView('AAMkAGI2TGuLAAA=', NOW - 86_400_000, NOW + 86_400_000);
  // Tomorrow's lunch starts after the window.
  assert.equal(events.length, 5);
  const request = microsoft.requests.at(-1);
  assert.equal(request.headers.prefer, 'outlook.timezone="UTC"');
  const url = new URL(request.url);
  assert.equal(url.pathname, '/v1.0/me/calendars/AAMkAGI2TGuLAAA%3D/calendarView');
  assert.equal(url.searchParams.get('startDateTime'), '2026-10-11T06:00:00.000Z');
  assert.match(url.searchParams.get('$select'), /showAs/);
});

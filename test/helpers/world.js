// -----------------------------------------------------------------------------
// A test world: the SDK's fake Gladys, a fake Microsoft (identity platform and
// Graph) answered through the `fetch` the integration is given, a settable
// clock and an in-memory store. No network. Times are read in Europe/Paris,
// the TZ the supervisor injects for a French instance.
// -----------------------------------------------------------------------------

import { readFileSync } from 'node:fs';
import { createFakeGladys } from '@gladysassistant/integration-sdk/testing';
import { registerHandlers } from '../../src/app.js';
import { createMemoryStore } from '../../src/store.js';

process.env.TZ = 'Europe/Paris';
// The integration logs its life cycle; LOG_LEVEL=debug npm test shows it.
process.env.LOG_LEVEL ??= 'silent';

export const NOW = Date.parse('2026-10-12T06:00:00.000Z'); // Monday 12 October 2026, 08:00 in Paris
export const CLIENT_ID = '0f1e2d3c-4b5a-4968-8776-655443322110';
export const ACCOUNT = { id: '00000000-0000-0000-0000-000000000001', name: 'Camille Martin' };
export const REFRESH_TOKEN = 'M.C5_BAY.0.U.-Cg9refresh-token-invented-for-the-tests!';

/** A fixture, parsed. */
export function fixture(path) {
  return JSON.parse(readFileSync(new URL(`../fixtures/${path}`, import.meta.url), 'utf8'));
}

/**
 * Open and completed tasks of the default list, due dates written the way To
 * Do stores them for a Paris user: the local midnight, in UTC.
 */
export function defaultTasks() {
  // The midnight of Paris (UTC+2 in October) of `date`, in UTC: 22:00 the day before.
  const due = (date) => {
    const before = new Date(Date.parse(`${date}T00:00:00Z`) - 86_400_000).toISOString();
    return { dateTime: `${before.slice(0, 10)}T22:00:00.0000000`, timeZone: 'UTC' };
  };
  return [
    {
      id: 'task-bins',
      title: 'Sortir les poubelles',
      status: 'notStarted',
      dueDateTime: due('2026-10-09'),
      createdDateTime: '2026-10-01T08:00:00Z',
    },
    {
      id: 'task-plants',
      title: 'Arroser les plantes',
      status: 'inProgress',
      dueDateTime: due('2026-10-11'),
      createdDateTime: '2026-10-02T08:00:00Z',
    },
    {
      id: 'task-bread',
      title: 'Acheter du pain',
      status: 'notStarted',
      dueDateTime: due('2026-10-12'),
      createdDateTime: '2026-10-03T08:00:00Z',
    },
    {
      id: 'task-filter',
      title: 'Changer le filtre de la VMC',
      status: 'notStarted',
      createdDateTime: '2026-09-20T08:00:00Z',
    },
    {
      id: 'task-done',
      title: 'Réserver le restaurant',
      status: 'completed',
      dueDateTime: due('2026-10-05'),
      createdDateTime: '2026-09-28T08:00:00Z',
    },
  ];
}

const json = (status, body, headers = {}) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json', ...headers },
  });

/**
 * A fake Microsoft. `data` holds what the account contains; `login` drives
 * the identity platform; every request lands in `requests`.
 */
export function createMicrosoft({
  calendars = [
    ...fixture('graph/calendars.json').value,
    { id: 'AAMkAGI2TGuLAAB=family', name: 'Famille', canEdit: true, isDefaultCalendar: false },
  ],
  events = { 'AAMkAGI2TGuLAAA=': fixture('graph/calendar-view.json').value },
  lists = fixture('graph/todo-lists.json').value,
  tasks = { 'AAMkADIyAAAAABrJAAA=': defaultTasks() },
  now = () => NOW,
} = {}) {
  const microsoft = {
    data: { calendars, events, lists, tasks },
    requests: [],
    login: {
      // Answers of the token endpoint to a device code, in order (then 'ok').
      devicePolls: ['pending', 'pending'],
      deviceCodeError: null,
      refreshError: null,
      issued: 0,
    },
    accessToken: null,
    refreshToken: REFRESH_TOKEN,
    account: { ...ACCOUNT, displayName: ACCOUNT.name },
    // Graph answers forced for the next requests: [{ status, body, headers }].
    graphFailures: [],
    unreachable: false,
    pageSize: 200,
    nextId: 1,
  };

  function issueTokens() {
    microsoft.login.issued += 1;
    microsoft.accessToken = `eyJ0eXAiOiJKV1Q.access-${microsoft.login.issued}`;
    microsoft.refreshToken = `${REFRESH_TOKEN}-${microsoft.login.issued}`;
    return {
      token_type: 'Bearer',
      scope: 'Calendars.ReadWrite Tasks.ReadWrite User.Read profile openid email',
      expires_in: 3599,
      ext_expires_in: 3599,
      access_token: microsoft.accessToken,
      refresh_token: microsoft.refreshToken,
    };
  }

  function login(url, params) {
    const endpoint = url.pathname.split('/').pop();
    if (endpoint === 'devicecode') {
      if (microsoft.login.deviceCodeError) {
        return json(400, microsoft.login.deviceCodeError);
      }
      return json(200, fixture('login/devicecode.json'));
    }
    if (params.get('grant_type') === 'refresh_token') {
      if (microsoft.login.refreshError) {
        return json(400, microsoft.login.refreshError);
      }
      if (params.get('refresh_token') !== microsoft.refreshToken) {
        return json(400, fixture('login/invalid-refresh-token.json'));
      }
      return json(200, issueTokens());
    }
    // Device code grant.
    if (params.get('device_code') !== fixture('login/devicecode.json').device_code) {
      return json(400, fixture('login/bad-device-code.json'));
    }
    const next = microsoft.login.devicePolls.shift() ?? 'ok';
    if (next === 'pending') {
      return json(400, fixture('login/authorization-pending.json'));
    }
    if (next !== 'ok') {
      return json(400, { error: next, error_codes: [70000] });
    }
    return json(200, issueTokens());
  }

  function page(items, url) {
    const skip = Number(url.searchParams.get('$skip') ?? 0);
    const size = Math.min(Number(url.searchParams.get('$top') ?? 100), microsoft.pageSize);
    const value = items.slice(skip, skip + size);
    const body = { value };
    if (skip + size < items.length) {
      const next = new URL(url);
      next.searchParams.set('$skip', String(skip + size));
      body['@odata.nextLink'] = next.href;
    }
    return json(200, body);
  }

  function graph(method, url, body) {
    const path = decodeURIComponent(url.pathname.replace(/^\/v1\.0/, ''));
    const parts = path.split('/').filter(Boolean);
    if (method === 'GET' && path === '/me') {
      return json(200, { id: microsoft.account.id, displayName: microsoft.account.displayName });
    }
    if (method === 'GET' && path === '/me/calendars') {
      return page(microsoft.data.calendars, url);
    }
    if (parts[1] === 'calendars' && parts[3] === 'calendarView' && method === 'GET') {
      const calendarId = parts[2];
      if (!microsoft.data.calendars.some((calendar) => calendar.id === calendarId)) {
        return json(404, { error: { code: 'ErrorItemNotFound', message: 'Not found.' } });
      }
      const from = Date.parse(url.searchParams.get('startDateTime'));
      const to = Date.parse(url.searchParams.get('endDateTime'));
      const overlapping = (microsoft.data.events[calendarId] ?? []).filter((event) => {
        const start = Date.parse(`${event.start.dateTime.slice(0, 19)}Z`);
        const end = Date.parse(`${event.end.dateTime.slice(0, 19)}Z`);
        return start < to && end > from;
      });
      return page(overlapping, url);
    }
    if (parts[1] === 'calendars' && parts[3] === 'events' && method === 'POST') {
      const event = {
        id: `created-event-${microsoft.nextId++}`,
        isAllDay: false,
        isCancelled: false,
        responseStatus: { response: 'organizer' },
        location: { displayName: '' },
        ...body,
      };
      (microsoft.data.events[parts[2]] ??= []).push(event);
      return json(201, event);
    }
    if (method === 'GET' && path === '/me/todo/lists') {
      return page(microsoft.data.lists, url);
    }
    if (parts[1] === 'todo' && parts[4] === 'tasks') {
      const listId = parts[3];
      if (!microsoft.data.lists.some((list) => list.id === listId)) {
        return json(404, { error: { code: 'ErrorItemNotFound', message: 'Not found.' } });
      }
      const listTasks = (microsoft.data.tasks[listId] ??= []);
      if (method === 'GET' && parts.length === 5) {
        const filter = url.searchParams.get('$filter');
        const open =
          filter === "status ne 'completed'"
            ? listTasks.filter((task) => task.status !== 'completed')
            : listTasks;
        return page(open, url);
      }
      if (method === 'POST' && parts.length === 5) {
        const task = {
          id: `created-task-${microsoft.nextId++}`,
          status: 'notStarted',
          importance: 'normal',
          createdDateTime: new Date(now()).toISOString(),
          ...body,
        };
        listTasks.push(task);
        return json(201, task);
      }
      if (method === 'PATCH' && parts.length === 6) {
        const task = listTasks.find((entry) => entry.id === parts[5]);
        if (!task) {
          return json(404, { error: { code: 'ErrorItemNotFound', message: 'Not found.' } });
        }
        Object.assign(task, body);
        return json(200, task);
      }
    }
    return json(400, { error: { code: 'BadRequest', message: `Unexpected ${method} ${path}` } });
  }

  microsoft.fetch = async (input, init = {}) => {
    const url = new URL(String(input));
    const method = init.method ?? 'GET';
    const headers = Object.fromEntries(
      Object.entries(init.headers ?? {}).map(([key, value]) => [key.toLowerCase(), value]),
    );
    const record = { method, url: url.href, host: url.host, headers, body: init.body };
    microsoft.requests.push(record);
    if (microsoft.unreachable) {
      throw new TypeError(`fetch failed (${url.href})`);
    }
    if (url.host === 'login.microsoftonline.com') {
      return login(url, new URLSearchParams(String(init.body ?? '')));
    }
    if (url.host !== 'graph.microsoft.com') {
      throw new Error(`Unexpected host ${url.host}`);
    }
    if (headers.authorization !== `Bearer ${microsoft.accessToken}`) {
      return json(401, { error: { code: 'InvalidAuthenticationToken', message: 'Expired.' } });
    }
    const failure = microsoft.graphFailures.shift();
    if (failure) {
      return json(failure.status, failure.body ?? {}, failure.headers);
    }
    return graph(method, url, init.body ? JSON.parse(init.body) : undefined);
  };

  /** The Graph requests (no login), as "METHOD /path". */
  microsoft.graphCalls = () =>
    microsoft.requests
      .filter((request) => request.host === 'graph.microsoft.com')
      .map(
        (request) =>
          `${request.method} ${decodeURIComponent(new URL(request.url).pathname.replace('/v1.0', ''))}`,
      );

  return microsoft;
}

/** Let the background work (sign-in wait, handlers) run. */
export const settle = () => new Promise((resolve) => setImmediate(resolve));

/**
 * @param {object} [options]
 * @param {boolean} [options.signedIn] a refresh token is already in /data
 * @param {object} [options.config] raw integration config
 * @param {object} [options.microsoft] createMicrosoft options
 * @param {'all' | 'none' | string[]} [options.add] devices the user adds (by name prefix)
 * @param {boolean} [options.timers]
 */
export async function createWorld({
  signedIn = true,
  config = {},
  microsoft: microsoftOptions = {},
  add = 'all',
  timers = false,
  connect = true,
  store,
} = {}) {
  const clock = { now: NOW };
  const now = () => clock.now;
  const microsoft = createMicrosoft({ now, ...microsoftOptions });
  const memoryStore =
    store ??
    createMemoryStore({
      tokens: signedIn
        ? { client_id: CLIENT_ID, tenant: 'common', refresh_token: REFRESH_TOKEN, account: ACCOUNT }
        : null,
    });
  const gladys = createFakeGladys({ config: { client_id: CLIENT_ID, ...config } });
  const app = registerHandlers(gladys, {
    store: memoryStore,
    fetch: microsoft.fetch,
    now,
    timers,
    // The sign-in wait and Graph's Retry-After advance the clock instead of sleeping.
    sleep: async (ms) => {
      clock.now += ms;
    },
  });
  const world = {
    gladys,
    app,
    monitor: app.monitor,
    microsoft,
    clock,
    store: memoryStore,

    /** The user adds the discovered devices whose name starts with one of `names`. */
    async addDevices(names = 'all') {
      for (const device of gladys.fake.discoveredDevices) {
        if (names === 'all' || names.some((name) => device.name.startsWith(name))) {
          await gladys.fake.deviceCreated(device);
        }
      }
    },

    /** The discovered device named `prefix…`. */
    device(prefix) {
      const found = gladys.fake.discoveredDevices.find((device) => device.name.startsWith(prefix));
      if (!found) {
        throw new Error(`No discovered device "${prefix}"`);
      }
      return found;
    },

    /** The last state published for a feature of a device. */
    state(prefix, feature) {
      return gladys.fake.lastState(`${world.device(prefix).external_id}:${feature}`);
    },

    /** Move the clock and run one timer step. */
    async advance(ms) {
      clock.now += ms;
      await app.monitor.tick();
    },

    /** Run the 30-second timer steps up to `ms` later. */
    async run(ms, step = 30_000) {
      for (let elapsed = 0; elapsed < ms; elapsed += step) {
        await world.advance(step);
      }
    },
  };
  if (connect) {
    await gladys.connect();
    if (add !== 'none') {
      await world.addDevices(add);
    }
  }
  return world;
}

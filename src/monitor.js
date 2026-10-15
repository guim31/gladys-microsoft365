// -----------------------------------------------------------------------------
// The monitor: reads Microsoft Graph for the devices the user added, turns it
// into device states and scene events, and runs the sign-in.
//
// Refresh: the core cannot be trusted to poll external devices
// (GladysAssistant/Gladys#3198: a device created without `should_poll` is
// never polled, and the 5.2.1 does not fix it), so the container runs its own
// timer: every 30 seconds it recomputes the states and the trigger instants
// from what it already read (no Graph call), and every `refresh_minutes` it
// reads Graph again. The core's onPoll stays as a fallback: it reads Graph
// only when the timer has not done it for a whole period.
//
// Graph budget, per added calendar or list: one request per period (plus one
// per 200 occurrences or 100 open tasks). Outlook allows 10,000 requests per
// 10 minutes per application and mailbox: a 1-minute period with 10 devices
// is 100.
//
// Logs: nothing at the info level for a normal cycle. Info keeps the life
// cycle (start, sign-in, sign-out, a session lost, Graph lost and back) and
// one line per user action; a failure that repeats is logged once.
// -----------------------------------------------------------------------------

import { createLogger } from '@gladysassistant/integration-sdk';
import { AuthError } from './auth.js';
import { MAX_LATENESS_MS, calendarState, normalizeEvents, transitions } from './calendar.js';
import { refreshMs } from './config.js';
import { KINDS, buildDevice, graphIdParam, parseDeviceId, platformIdOf } from './devices.js';
import { formatStart, truncate } from './format.js';
import { GraphError, needsSignIn } from './graph.js';
import { TEXT, everyLanguage, pick } from './i18n.js';
import { createIssueLog } from './issues.js';
import { eventTriggerData, overdueTriggerData } from './scenes.js';
import { DAY_MS, HOUR_MS, MINUTE_MS, localDateKey } from './time.js';
import { listState, normalizeTasks } from './todo.js';

const logger = createLogger({ name: 'monitor' });

export const TICK_MS = 30_000;
export const AGENDA_WIDGET = 'agenda';
// Occurrences read: from yesterday (an event that just ended still fires its
// "ended" trigger) to two weeks ahead (the "next event").
const PAST_MS = DAY_MS;
const AHEAD_MS = 14 * DAY_MS;
// Unchanged values are published again this often: the dashboard shows "no
// recent value" past 48 hours.
const HEARTBEAT_MS = 6 * HOUR_MS;
// Data older than this does not fire triggers (Graph unreachable for hours).
const STALE_MS = 2 * HOUR_MS;
// At most this many "task overdue" events per list and check.
const MAX_OVERDUE_EVENTS = 20;

const defaultSleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * @param {object} options
 * @param {import('@gladysassistant/integration-sdk').GladysIntegration} options.gladys
 * @param {ReturnType<import('./store.js').createStore>} options.store
 * @param {ReturnType<import('./auth.js').createAuth>} options.auth
 * @param {ReturnType<import('./graph.js').createGraphClient>} options.graph
 * @param {() => object} options.getConfig
 * @param {() => number} [options.now]
 * @param {boolean} [options.timers] run the 30-second timer (tests drive it by hand)
 * @param {(ms: number) => Promise<void>} [options.sleep] wait between two sign-in checks
 */
export function createMonitor({
  gladys,
  store,
  auth,
  graph,
  getConfig,
  now = Date.now,
  timers = true,
  sleep = defaultSleep,
}) {
  const issues = createIssueLog(logger);
  /** What was read, by device external id: `{ kind, events | tasks, syncedAt }`. */
  const cache = new Map();
  let lastSyncAt = 0;
  let lastCheckAt;
  let lastStatus = null;
  let widgetSignature = null;
  let pendingSignIn = null;
  let signInRun = null;
  // The refresh token was refused: say "expired", not just "not signed in".
  let sessionLost = false;
  let timer = null;
  let stopped = false;
  let chain = Promise.resolve();

  /** Runs `task` after the previous one: syncs, checks and actions never overlap. */
  function exclusive(task) {
    const run = chain.then(task);
    chain = run.catch(() => {});
    return run;
  }

  // --- Connection status (Configuration screen) -----------------------------

  async function setStatus(connected, messageKey, ...args) {
    const message = messageKey ? everyLanguage(messageKey, ...args) : undefined;
    const signature = JSON.stringify([connected, message]);
    if (signature === lastStatus || gladys.connected === false) {
      return;
    }
    try {
      await gladys.setConnectionStatus(connected, message);
      lastStatus = signature;
    } catch (err) {
      logger.debug(`Connection status not sent (${err.message})`);
    }
  }

  /** The status when nothing is wrong with Graph: signed in, or what is missing. */
  function idleStatus() {
    const config = getConfig();
    if (!config.client_id) {
      return setStatus(false, 'missingClientId');
    }
    if (pendingSignIn) {
      return setStatus(
        false,
        'signInPending',
        pendingSignIn.verificationUri,
        pendingSignIn.userCode,
      );
    }
    if (!auth.signedIn(config)) {
      return setStatus(false, sessionLost ? 'sessionExpired' : 'notSignedIn');
    }
    return setStatus(true);
  }

  // --- The devices the user added -------------------------------------------

  /** Our created devices: `{ device, kind, platformId, graphId }`. */
  function createdDevices() {
    const found = [];
    for (const device of gladys.devices ?? []) {
      const parsed = parseDeviceId(gladys, device.external_id);
      if (parsed) {
        const graphId = graphIdParam(device) ?? store.graphId(parsed.platformId);
        found.push({ device, ...parsed, graphId });
      }
    }
    return found;
  }

  function findCreated(externalId, kind) {
    return createdDevices().find(
      (entry) => entry.device.external_id === externalId && (!kind || entry.kind === kind),
    );
  }

  // --- Reading Graph --------------------------------------------------------

  /** A Graph or sign-in failure: what to tell, and whether to stop the cycle. */
  async function handleFailure(err, key, what) {
    if (needsSignIn(err)) {
      if (err.code !== 'not_signed_in' && !sessionLost) {
        sessionLost = true;
        logger.warn(`Microsoft session lost (${err.message}): sign in again`);
      }
      await idleStatus();
      return 'stop';
    }
    if (err instanceof GraphError && err.throttled) {
      const minutes = Math.max(1, Math.ceil((err.retryAt - now()) / MINUTE_MS));
      issues.fail('throttled', `Microsoft Graph throttles the integration for ${minutes} min`);
      return 'stop';
    }
    const unreachable =
      (err instanceof GraphError && (err.status === 0 || err.status >= 500)) ||
      (err instanceof AuthError && ['network_error', 'timeout'].includes(err.code)) ||
      (err instanceof AuthError && err.status >= 500);
    if (unreachable) {
      issues.fail('graph', `Microsoft 365 unreachable (${err.message})`);
      await setStatus(false, 'unreachable');
      return 'stop';
    }
    issues.fail(key, `${what}: ${err.message}`);
    return 'continue';
  }

  async function readDevice(entry) {
    const key = `device:${entry.device.external_id}`;
    if (!entry.graphId) {
      issues.fail(key, `${entry.device.name}: unknown to the integration, run a discovery`);
      return;
    }
    const at = now();
    if (entry.kind === KINDS.CALENDAR) {
      const raw = await graph.calendarView(entry.graphId, at - PAST_MS, at + AHEAD_MS);
      cache.set(entry.device.external_id, {
        kind: entry.kind,
        events: normalizeEvents(raw),
        syncedAt: at,
      });
    } else {
      const raw = await graph.openTasks(entry.graphId);
      cache.set(entry.device.external_id, {
        kind: entry.kind,
        tasks: normalizeTasks(raw),
        syncedAt: at,
      });
    }
    issues.ok(key, `${entry.device.name}: read again`);
  }

  /** Read Graph for the given devices (all the added ones by default). */
  async function readGraph(entries = createdDevices()) {
    const config = getConfig();
    if (!auth.signedIn(config)) {
      return;
    }
    if (graph.blockedUntil) {
      logger.debug('Microsoft Graph throttling in progress: read skipped');
      return;
    }
    let reached = entries.length === 0;
    for (const entry of entries) {
      try {
        await readDevice(entry);
        reached = true;
      } catch (err) {
        if (err instanceof GraphError && err.status === 404) {
          // Deleted in Outlook or To Do: nothing to read any more.
          cache.delete(entry.device.external_id);
          issues.fail(
            `device:${entry.device.external_id}`,
            `${entry.device.name}: no longer exists in Microsoft 365`,
          );
          reached = true;
          continue;
        }
        const outcome = await handleFailure(
          err,
          `device:${entry.device.external_id}`,
          entry.device.name,
        );
        if (outcome === 'stop') {
          return;
        }
        reached = true;
      }
    }
    if (reached) {
      issues.ok('graph', 'Microsoft 365 reachable again');
      issues.ok('throttled', 'Microsoft Graph throttling over');
      await idleStatus();
    }
  }

  // --- States, triggers, widget ---------------------------------------------

  function calendarStates(entry, data, at, language) {
    const ids = gladys.externalIds(entry.kind, entry.platformId);
    const state = calendarState(data.events, at);
    const none = pick(TEXT, language).none;
    const nextText = state.next
      ? `${truncate(state.next.title || none, 80)} · ${formatStart(state.next, language, at)}`
      : none;
    return [
      { device_feature_external_id: ids.feature('busy'), state: state.busy ? 1 : 0 },
      {
        device_feature_external_id: ids.feature('out_of_office'),
        state: state.outOfOffice ? 1 : 0,
      },
      {
        device_feature_external_id: ids.feature('current_event'),
        text: truncate(state.current?.title || (state.current ? '…' : none), 100),
      },
      { device_feature_external_id: ids.feature('next_event'), text: truncate(nextText, 120) },
    ];
  }

  function listStates(entry, state) {
    const ids = gladys.externalIds(entry.kind, entry.platformId);
    return [
      { device_feature_external_id: ids.feature('open_tasks'), state: state.open },
      { device_feature_external_id: ids.feature('overdue_tasks'), state: state.overdue },
      { device_feature_external_id: ids.feature('due_today_tasks'), state: state.dueToday },
    ];
  }

  async function fire(key, data) {
    try {
      await gladys.publishSceneEvent(key, data);
      issues.ok('scene-event', 'Scene events accepted again');
    } catch (err) {
      issues.fail('scene-event', `Scene event ${key} refused (${err.message})`);
    }
  }

  /** States, trigger instants and overdue tasks at `now`, from what was read. */
  async function check() {
    if (gladys.connected === false) {
      return;
    }
    const config = getConfig();
    const at = now();
    const today = localDateKey(at);
    const from = lastCheckAt === undefined ? at : Math.max(lastCheckAt, at - MAX_LATENESS_MS);
    lastCheckAt = at;
    const states = [];
    const events = [];
    let storeChanged = false;

    for (const entry of createdDevices()) {
      const data = cache.get(entry.device.external_id);
      if (!data || data.kind !== entry.kind) {
        continue;
      }
      const fresh = at - data.syncedAt <= STALE_MS;
      if (entry.kind === KINDS.CALENDAR) {
        states.push(...calendarStates(entry, data, at, config.language));
        if (fresh) {
          for (const transition of transitions(data.events, from, at)) {
            events.push([
              transition.key,
              eventTriggerData(transition, entry.device, config.trigger_keywords, config.language),
            ]);
          }
        }
        continue;
      }
      const state = listState(data.tasks, today);
      states.push(...listStates(entry, state));
      if (!fresh) {
        continue;
      }
      const overdueIds = state.overdueTasks.map((task) => task.id);
      const known = store.overdue(entry.platformId);
      if (known === undefined) {
        // First look at this list: what is already overdue is not news.
        store.setOverdue(entry.platformId, overdueIds);
        storeChanged = true;
        continue;
      }
      const fresher = state.overdueTasks.filter((task) => !known.includes(task.id));
      for (const task of fresher.slice(0, MAX_OVERDUE_EVENTS)) {
        events.push(['task_overdue', overdueTriggerData(task, entry.device, config.language, at)]);
      }
      if (JSON.stringify(known) !== JSON.stringify(overdueIds)) {
        store.setOverdue(entry.platformId, overdueIds);
        storeChanged = true;
      }
    }

    if (states.length > 0) {
      try {
        await gladys.publishChangedStates(states, { heartbeat: HEARTBEAT_MS });
        issues.ok('states', 'States accepted again by Gladys');
      } catch (err) {
        issues.fail('states', `States refused by Gladys (${err.message})`);
      }
    }
    for (const [key, data] of events) {
      logger.debug(`Scene event ${key}`);
      await fire(key, data);
    }
    if (storeChanged) {
      await store.save();
    }
    refreshWidgetIfChanged(at);
  }

  /** Nudge the dashboards when what the widget shows changed. */
  function refreshWidgetIfChanged(at) {
    const today = localDateKey(at);
    const signature = JSON.stringify(
      [...cache.entries()].map(([id, data]) => [
        id,
        data.kind === KINDS.CALENDAR
          ? data.events
              .filter((event) => event.endMs > at && localDateKey(event.startMs) <= today)
              .map((event) => [event.id, event.title, event.startMs, event.endMs])
          : data.tasks.map((task) => [task.id, task.title, task.dueDate]),
      ]),
    );
    if (widgetSignature !== null && signature !== widgetSignature) {
      gladys.requestWidgetRefresh(AGENDA_WIDGET);
    }
    widgetSignature = signature;
  }

  // --- Discovery ------------------------------------------------------------

  async function ensureAccount(config) {
    const known = auth.account(config);
    if (known?.id) {
      return known;
    }
    const me = await graph.me();
    const account = { id: String(me?.id ?? ''), name: String(me?.displayName ?? '') };
    if (!account.id) {
      throw new GraphError(200, 'no_account_id');
    }
    await auth.setAccount(config, account);
    return account;
  }

  /** Publish one device per calendar and per To Do list of the account. */
  async function discover() {
    const config = getConfig();
    if (!auth.signedIn(config)) {
      await gladys.publishDiscoveredDevices([]);
      return null;
    }
    const account = await ensureAccount(config);
    const calendars = await graph.calendars();
    const lists = await graph.todoLists();
    const sources = [
      ...calendars.map((calendar) => ({ kind: KINDS.CALENDAR, ...calendar })),
      ...lists.map((list) => ({ kind: KINDS.TODO_LIST, name: list.displayName, ...list })),
    ]
      .filter((source) => typeof source.id === 'string' && source.id)
      .map((source) => ({
        kind: source.kind,
        platformId: platformIdOf(account.id, source.id),
        graphId: source.id,
        name: String(source.name ?? '').trim(),
      }));
    store.setGraphIds(
      Object.fromEntries(sources.map((source) => [source.platformId, source.graphId])),
    );
    await store.save();
    await gladys.publishDiscoveredDevices(
      sources.map((source) => buildDevice(gladys, source, config.language)),
    );
    logger.debug(`Discovery: ${calendars.length} calendar(s), ${lists.length} To Do list(s)`);
    return { account, calendars: calendars.length, lists: lists.length };
  }

  // --- Sign-in --------------------------------------------------------------

  async function finishSignIn(attempt, tokens) {
    const { config } = attempt;
    await auth.adopt(config, tokens, null);
    pendingSignIn = null;
    sessionLost = false;
    let name = '';
    try {
      const account = await ensureAccount(config);
      name = account.name;
    } catch (err) {
      logger.warn(`Signed in, but the account could not be read (${err.message})`);
    }
    logger.info('Signed in to Microsoft 365');
    await setStatus(true);
    await exclusive(async () => {
      try {
        await discover();
      } catch (err) {
        await handleFailure(err, 'discovery', 'Discovery');
      }
      await readGraph();
      lastSyncAt = now();
      await check();
    });
    return name;
  }

  /**
   * Waits for the user to type the code, until it expires. A lost Gladys
   * connection does not stop it: the user is typing the code meanwhile.
   */
  async function waitForSignIn(attempt) {
    let interval = attempt.intervalMs;
    while (pendingSignIn === attempt && now() < attempt.expiresAt) {
      await sleep(interval);
      if (pendingSignIn !== attempt) {
        return;
      }
      let answer;
      try {
        answer = await auth.pollDeviceCode(attempt.config, attempt.deviceCode);
      } catch (err) {
        if (err.code === 'network_error' || err.code === 'timeout' || err.status >= 500) {
          logger.debug(`Sign-in check failed (${err.message}), retrying`);
          continue;
        }
        pendingSignIn = null;
        const declined = err.code === 'authorization_declined';
        const expired = err.code === 'expired_token' || err.code === 'bad_verification_code';
        logger.info(
          `Sign-in ${declined ? 'declined' : expired ? 'expired' : `refused (${err.message})`}`,
        );
        await setStatus(
          false,
          declined ? 'signInDeclined' : expired ? 'signInExpired' : 'signInRefused',
          err.message,
        );
        return;
      }
      if (answer.status === 'slow_down') {
        interval += 5000;
      } else if (answer.status === 'ok') {
        await finishSignIn(attempt, answer.tokens);
        return;
      }
    }
    if (pendingSignIn === attempt) {
      pendingSignIn = null;
      logger.info('Sign-in code expired');
      await setStatus(false, 'signInExpired');
    }
  }

  // --- Timer ----------------------------------------------------------------

  /** One timer step: read Graph when a period has passed, check in any case. */
  function tick() {
    return exclusive(async () => {
      const config = getConfig();
      if (now() - lastSyncAt >= refreshMs(config)) {
        lastSyncAt = now();
        await readGraph();
      }
      await check();
    });
  }

  function schedule() {
    clearTimeout(timer);
    timer = null;
    if (!timers || stopped) {
      return;
    }
    timer = setTimeout(async () => {
      timer = null;
      try {
        await tick();
      } catch (err) {
        logger.error(`Refresh cycle failed (${err.message})`);
      } finally {
        schedule();
      }
    }, TICK_MS);
  }

  return {
    cache,
    tick,
    /** Resolves once the work queued so far is done (tests). */
    idle: () => chain,
    discover: () => exclusive(discover),
    idleStatus,

    /** (Re)start after a connection or a new config: status, discovery, read. */
    async start() {
      stopped = false;
      schedule();
      await idleStatus();
      await exclusive(async () => {
        if (auth.signedIn(getConfig())) {
          try {
            await discover();
          } catch (err) {
            await handleFailure(err, 'discovery', 'Discovery');
          }
        }
        lastSyncAt = now();
        await readGraph();
        await check();
      });
    },

    stop() {
      stopped = true;
      clearTimeout(timer);
      timer = null;
    },

    /** The configuration changed: a new application or tenant ends the sign-in. */
    async configChanged(previous) {
      const config = getConfig();
      if (previous.client_id !== config.client_id || previous.tenant !== config.tenant) {
        pendingSignIn = null;
        cache.clear();
      }
      lastSyncAt = 0;
      await this.start();
    },

    /** The core's onPoll: reads Graph only if the timer has not for a period. */
    pollFromCore(device) {
      return exclusive(async () => {
        const entry = findCreated(device.external_id);
        if (!entry || now() - lastSyncAt < refreshMs(getConfig())) {
          return;
        }
        lastSyncAt = now();
        await readGraph();
        await check();
      });
    },

    /** A device was added or updated: read it and publish everything now. */
    deviceChanged(device) {
      return exclusive(async () => {
        const entry = findCreated(device.external_id);
        if (!entry) {
          return;
        }
        await readGraph([entry]);
        await check();
        gladys.requestWidgetRefresh(AGENDA_WIDGET);
      });
    },

    deviceDeleted(device) {
      cache.delete(device.external_id);
      issues.forget(`device:${device.external_id}`);
    },

    /** Re-read one device after a change made by a scene or the widget. */
    refreshDevice(externalId) {
      return exclusive(async () => {
        const entry = findCreated(externalId);
        if (entry) {
          await readGraph([entry]);
          await check();
        }
      });
    },

    /**
     * Start the sign-in (action "Sign in"): resolves the message with the
     * code at once; the wait for the user goes on in the background.
     */
    async startSignIn() {
      const config = getConfig();
      if (!config.client_id) {
        return everyLanguage('missingClientId');
      }
      const minutesLeft = (attempt) =>
        Math.max(1, Math.floor((attempt.expiresAt - now()) / 60_000));
      if (
        pendingSignIn &&
        pendingSignIn.config.client_id === config.client_id &&
        pendingSignIn.expiresAt - now() > MINUTE_MS
      ) {
        // A second click: the same code is still valid.
        return everyLanguage(
          'signInCode',
          pendingSignIn.verificationUri,
          pendingSignIn.userCode,
          minutesLeft(pendingSignIn),
        );
      }
      let code;
      try {
        code = await auth.requestDeviceCode(config);
      } catch (err) {
        logger.info(`Sign-in refused by Microsoft (${err.message})`);
        return everyLanguage('signInRefused', err.message);
      }
      const attempt = { ...code, config };
      pendingSignIn = attempt;
      logger.info('Sign-in started: waiting for the code to be entered at Microsoft');
      await setStatus(false, 'signInPending', attempt.verificationUri, attempt.userCode);
      const message = everyLanguage(
        'signInCode',
        attempt.verificationUri,
        attempt.userCode,
        minutesLeft(attempt),
      );
      signInRun = waitForSignIn(attempt).catch((err) => {
        logger.error(`Sign-in failed (${err.message})`);
      });
      return message;
    },

    /** The background wait of the last sign-in (tests). */
    get signInDone() {
      return signInRun ?? Promise.resolve();
    },

    get pendingSignIn() {
      return pendingSignIn;
    },

    async signOut() {
      pendingSignIn = null;
      sessionLost = false;
      await auth.signOut();
      cache.clear();
      logger.info('Signed out of Microsoft 365');
      await setStatus(false, 'notSignedIn');
      try {
        await gladys.publishDiscoveredDevices([]);
      } catch (err) {
        logger.debug(`Discovered devices not cleared (${err.message})`);
      }
      return everyLanguage('signedOut');
    },

    /** Action "Check the connection": the account and what it holds. */
    async checkConnection() {
      const config = getConfig();
      if (!config.client_id) {
        return everyLanguage('missingClientId');
      }
      if (!auth.signedIn(config)) {
        return everyLanguage(
          pendingSignIn ? 'signInPending' : 'notSignedIn',
          pendingSignIn?.verificationUri,
          pendingSignIn?.userCode,
        );
      }
      try {
        const found = await exclusive(discover);
        await idleStatus();
        return everyLanguage('report', found.account.name || '?', found.calendars, found.lists);
      } catch (err) {
        await handleFailure(err, 'discovery', 'Check');
        if (needsSignIn(err)) {
          return everyLanguage('sessionExpired');
        }
        return everyLanguage('failed', err.message);
      }
    },

    /** For the widget and the scene actions. */
    findCreated,
    createdDevices,
    get signedIn() {
      return auth.signedIn(getConfig());
    },
  };
}

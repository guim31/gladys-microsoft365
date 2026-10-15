// -----------------------------------------------------------------------------
// Every SDK handler, registered before connect(). index.js only builds the
// real client; the tests pass the SDK's fake Gladys and a fake Microsoft.
// -----------------------------------------------------------------------------

import { createLogger } from '@gladysassistant/integration-sdk';
import { createAuth } from './auth.js';
import { normalizeConfig } from './config.js';
import { createGraphClient } from './graph.js';
import { createMonitor } from './monitor.js';
import { SCENE_ACTIONS } from './scenes.js';
import { createStore } from './store.js';
import { AGENDA_WIDGET, agendaAction, agendaContent } from './widgets.js';

const logger = createLogger({ name: 'app' });

/** The manifest `actions`, by key. */
export const ACTIONS = {
  sign_in: (monitor) => monitor.startSignIn(),
  sign_out: (monitor) => monitor.signOut(),
  check_connection: (monitor) => monitor.checkConnection(),
};

/**
 * @param {import('@gladysassistant/integration-sdk').GladysIntegration} gladys
 * @param {object} [options]
 * @param {ReturnType<typeof createStore>} [options.store]
 * @param {typeof fetch} [options.fetch] the way to Microsoft (tests: a fake)
 * @param {() => number} [options.now]
 * @param {boolean} [options.timers] run the refresh timer
 * @param {(ms: number) => Promise<void>} [options.sleep]
 */
export function registerHandlers(
  gladys,
  { store = createStore(), fetch = globalThis.fetch, now = Date.now, timers = true, sleep } = {},
) {
  let config = normalizeConfig(gladys.config ?? {});
  const getConfig = () => config;
  const auth = createAuth({ store, fetch, now });
  const graph = createGraphClient({ auth, getConfig, fetch, now, sleep });
  const monitor = createMonitor({ gladys, store, auth, graph, getConfig, now, timers, sleep });
  const context = {
    monitor,
    graph,
    gladys,
    now,
    get config() {
      return config;
    },
  };
  const ready = store.load();

  gladys.onScanRequest(async () => {
    await ready;
    try {
      await monitor.discover();
    } catch (err) {
      logger.warn(`Discovery failed (${err.message})`);
    }
  });

  // Fallback only: the monitor refreshes the devices on its own timer.
  gladys.onPoll(async (device) => {
    await ready;
    await monitor.pollFromCore(device);
  });
  // A device just added: Gladys dropped the states sent before, send them now.
  gladys.onDeviceCreated(async (device) => {
    await ready;
    await monitor.deviceChanged(device);
  });
  gladys.onDeviceUpdated(async (device) => {
    await ready;
    await monitor.deviceChanged(device);
  });
  gladys.onDeviceDeleted(async (device) => {
    monitor.deviceDeleted(device);
  });

  for (const [key, handler] of Object.entries(ACTIONS)) {
    gladys.onAction(key, async () => {
      await ready;
      logger.info(`Action ${key}`);
      return handler(monitor);
    });
  }

  for (const [key, handler] of Object.entries(SCENE_ACTIONS)) {
    gladys.onSceneAction(key, async (fields) => {
      await ready;
      return handler(context, fields);
    });
  }

  gladys.onWidgetGet(AGENDA_WIDGET, async (request) => {
    await ready;
    return agendaContent(context, { ...request, now: now() });
  });
  gladys.onWidgetAction(AGENDA_WIDGET, async (actionKey, params, { values } = {}) => {
    await ready;
    logger.info(`Widget action ${actionKey}`);
    return agendaAction(context, { actionKey, params, values });
  });

  gladys.onConfigUpdated(async (newConfig) => {
    await ready;
    const previous = config;
    config = normalizeConfig(newConfig);
    logger.info('Configuration updated');
    await monitor.configChanged(previous);
  });

  // On every (re)connection: the config, the status, the discovery, a read.
  gladys.on('connected', async () => {
    try {
      await ready;
      config = normalizeConfig(await gladys.getConfig());
      await monitor.start();
      // Open dashboards missed the nudges sent while the socket was down.
      gladys.requestWidgetRefresh(AGENDA_WIDGET);
    } catch (err) {
      logger.error(`Initialization failed (${err.message})`);
    }
  });
  gladys.on('disconnected', () => {
    monitor.stop();
  });

  return {
    monitor,
    auth,
    graph,
    ready,
    getConfig,
    stop() {
      monitor.stop();
    },
  };
}

// -----------------------------------------------------------------------------
// The devices Gladys discovers: one per Outlook calendar, one per To Do list.
// The user adds the ones they want; only those are read.
//
// External ids are forever: they derive from the Graph id of the calendar or
// list AND from the account (a digest of both, short and URL-safe), so
// signing in to another account gives other devices instead of feeding the
// old ones with someone else's data. The Graph id itself travels in the
// device params (GRAPH_ID), and in /data for the devices not yet added.
// -----------------------------------------------------------------------------

import { createHash } from 'node:crypto';
import { DEVICE_POLL_FREQUENCIES } from '@gladysassistant/integration-sdk';
import { NAMES, pick } from './i18n.js';
import { truncate } from './format.js';

export const KINDS = { CALENDAR: 'calendar', TODO_LIST: 'todo_list' };
export const GRAPH_ID_PARAM = 'GRAPH_ID';

const binary = (key) => ({
  key,
  category: 'input',
  type: 'binary',
  min: 0,
  max: 1,
  keep_history: true,
});
const text = (key) => ({
  key,
  category: 'text',
  type: 'text',
  min: 0,
  max: 0,
  keep_history: false,
});
const count = (key) => ({
  key,
  category: 'counter-sensor',
  type: 'integer',
  min: 0,
  max: 10000,
  keep_history: true,
});

export const FEATURES = {
  [KINDS.CALENDAR]: [
    binary('busy'),
    binary('out_of_office'),
    text('current_event'),
    text('next_event'),
  ],
  [KINDS.TODO_LIST]: [count('open_tasks'), count('overdue_tasks'), count('due_today_tasks')],
};

/** The stable platform id of a calendar or list of an account. */
export function platformIdOf(accountId, graphId) {
  return createHash('sha256').update(`${accountId}\n${graphId}`).digest('hex').slice(0, 20);
}

/**
 * A discovered device.
 * @param {import('@gladysassistant/integration-sdk').GladysIntegration} gladys
 * @param {{ kind: string, platformId: string, graphId: string, name: string }} source
 * @param {string} language of the feature names
 */
export function buildDevice(gladys, { kind, platformId, graphId, name }, language) {
  const ids = gladys.externalIds(kind, platformId);
  const names = pick(NAMES, language);
  const brand = kind === KINDS.CALENDAR ? 'Outlook' : 'To Do';
  return {
    name: `${truncate(name || brand, 60)} (${brand})`,
    external_id: ids.device,
    // The container refreshes the devices itself (src/monitor.js); the core's
    // poll is only a fallback, once a minute at most.
    should_poll: true,
    poll_frequency: DEVICE_POLL_FREQUENCIES.EVERY_MINUTES,
    params: [{ name: GRAPH_ID_PARAM, value: graphId }],
    features: FEATURES[kind].map(({ key, ...feature }) => ({
      ...feature,
      name: names[key],
      external_id: ids.feature(key),
      read_only: true,
      has_feedback: false,
    })),
  };
}

/**
 * The kind and platform id of one of our devices, from its external id.
 * @returns {{ kind: string, platformId: string } | null}
 */
export function parseDeviceId(gladys, externalId) {
  const prefix = gladys.externalId('');
  const rest = String(externalId ?? '').startsWith(prefix)
    ? String(externalId).slice(prefix.length)
    : '';
  const [kind, platformId, extra] = rest.split(':');
  if (!Object.values(KINDS).includes(kind) || !platformId || extra !== undefined) {
    return null;
  }
  return { kind, platformId };
}

/** The Graph id stored in the params of a created device. */
export function graphIdParam(device) {
  const param = (device?.params ?? []).find((entry) => entry?.name === GRAPH_ID_PARAM);
  return typeof param?.value === 'string' && param.value ? param.value : null;
}

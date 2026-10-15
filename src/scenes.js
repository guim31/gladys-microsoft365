// -----------------------------------------------------------------------------
// Scene triggers (fired by src/monitor.js) and scene actions (manifest
// `scene_triggers` and `scene_actions`). Keys are forever once published.
//
// The core only compares a trigger filter for equality: "the title contains
// a keyword" is decided here, against the keywords listed in the
// configuration, and the event carries the keyword it matched; the scene
// filters on that keyword.
// -----------------------------------------------------------------------------

import { createLogger } from '@gladysassistant/integration-sdk';
import { matchKeyword } from './calendar.js';
import { formatDueDate, formatStart, formatTime, truncate } from './format.js';
import { KINDS } from './devices.js';
import { localizedError } from './i18n.js';
import { MINUTE_MS, addDays, localDateKey } from './time.js';
import { findTask, normalizeTasks } from './todo.js';

const logger = createLogger({ name: 'scenes' });

export const SCENE_TRIGGER_KEYS = [
  'event_started',
  'event_ended',
  'event_upcoming',
  'task_overdue',
];

/** The data of an `event_started` / `event_ended` / `event_upcoming` event. */
export function eventTriggerData(transition, device, keywords, language) {
  const { event } = transition;
  const end = event.allDay ? '' : formatTime(event.endMs, language);
  return {
    calendar: device.external_id,
    calendar_name: truncate(device.name, 200),
    keyword: matchKeyword(event.title, keywords),
    title: truncate(event.title, 500),
    location: truncate(event.location, 500),
    start: event.allDay
      ? formatStart(event, language, event.startMs)
      : formatTime(event.startMs, language),
    end,
    ...(transition.key === 'event_upcoming' ? { minutes: transition.minutes } : {}),
  };
}

/** The data of a `task_overdue` event. */
export function overdueTriggerData(task, device, language, now) {
  return {
    list: device.external_id,
    list_name: truncate(device.name, 200),
    title: truncate(task.title, 500),
    due: formatDueDate(task.dueDate, language, now),
  };
}

/** The created device of a scene field, of the expected kind. */
function deviceOf(monitor, externalId, kind, language) {
  const entry = monitor.findCreated(externalId, kind);
  if (!entry?.graphId) {
    throw localizedError(language, kind === KINDS.CALENDAR ? 'unknownCalendar' : 'unknownList');
  }
  if (!monitor.signedIn) {
    throw localizedError(language, 'notSignedIn');
  }
  return entry;
}

function cleanTitle(value, language) {
  const title = String(value ?? '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 255);
  if (!title) {
    throw localizedError(language, 'emptyTitle');
  }
  return title;
}

/** Re-read the device after a change, without failing the action on a read error. */
async function refreshAfter(monitor, externalId) {
  try {
    await monitor.refreshDevice(externalId);
  } catch (err) {
    logger.debug(`Refresh after a scene action failed (${err.message})`);
  }
}

/**
 * Add a task to a To Do list.
 * @param {{ list: string, title: string, due_in_days?: number }} fields
 */
export async function addTask({ monitor, graph, config, now }, fields) {
  const entry = deviceOf(monitor, fields.list, KINDS.TODO_LIST, config.language);
  const title = cleanTitle(fields.title, config.language);
  const days = Number(fields.due_in_days);
  const dueDate =
    fields.due_in_days === undefined || fields.due_in_days === null || !Number.isFinite(days)
      ? undefined
      : addDays(localDateKey(now()), Math.min(365, Math.max(0, Math.round(days))));
  await graph.createTask(entry.graphId, { title, dueDate });
  logger.info(`Task added to a To Do list${dueDate ? ', with a due date' : ''}`);
  await refreshAfter(monitor, entry.device.external_id);
  return { title };
}

/**
 * Complete the open task of a list whose title is (or contains) a text.
 * @param {{ list: string, title: string }} fields
 */
export async function completeTask({ monitor, graph, config }, fields) {
  const entry = deviceOf(monitor, fields.list, KINDS.TODO_LIST, config.language);
  const wanted = cleanTitle(fields.title, config.language);
  // Read now: the cache may be minutes old, and the task added since.
  const tasks = normalizeTasks(await graph.openTasks(entry.graphId));
  const { task, reason } = findTask(tasks, wanted);
  if (!task) {
    logger.info(
      `No task completed: ${reason === 'ambiguous' ? 'several tasks match' : 'no open task matches'}`,
    );
    return { completed: 0, title: '' };
  }
  await graph.completeTask(entry.graphId, task.id);
  logger.info('Task completed in a To Do list');
  await refreshAfter(monitor, entry.device.external_id);
  return { completed: 1, title: task.title };
}

/**
 * Add an event to an Outlook calendar, starting in N minutes.
 * @param {{ calendar: string, title: string, start_in_minutes?: number, duration_minutes?: number }} fields
 */
export async function addEvent({ monitor, graph, config, now }, fields) {
  const entry = deviceOf(monitor, fields.calendar, KINDS.CALENDAR, config.language);
  const title = cleanTitle(fields.title, config.language);
  const startIn = Math.min(
    10080,
    Math.max(0, Math.round(Number(fields.start_in_minutes ?? 0) || 0)),
  );
  const duration = Math.min(
    1440,
    Math.max(5, Math.round(Number(fields.duration_minutes ?? 60) || 60)),
  );
  // Whole minutes: an event at 14:03:27 would read oddly in Outlook.
  const startMs = Math.ceil((now() + startIn * MINUTE_MS) / MINUTE_MS) * MINUTE_MS;
  await graph.createEvent(entry.graphId, {
    subject: title,
    startMs,
    endMs: startMs + duration * MINUTE_MS,
  });
  logger.info('Event added to an Outlook calendar');
  await refreshAfter(monitor, entry.device.external_id);
  return { start: formatTime(startMs, config.language) };
}

export const SCENE_ACTIONS = {
  add_task: addTask,
  complete_task: completeTask,
  add_event: addEvent,
};

// -----------------------------------------------------------------------------
// Dashboard widget "Today": what is left of today's Outlook agenda, and the
// open tasks of a To Do list, with a button to add a task (a form behind the
// tap: Gladys 5.2.0 and later).
//
// Built from what the monitor already read: displaying the widget never calls
// Microsoft Graph. The counts are live tiles bound to the list device.
// -----------------------------------------------------------------------------

import { WIDGET_COLORS } from '@gladysassistant/integration-sdk';
import { KINDS } from './devices.js';
import { formatDueDate, formatRange, truncate } from './format.js';
import { NAMES, TEXT, everyLanguage, pick, textInEveryLanguage } from './i18n.js';
import { AGENDA_WIDGET } from './monitor.js';
import { addTask } from './scenes.js';
import { localDateKey, localMidnight } from './time.js';

export { AGENDA_WIDGET };
export const WIDGET_ACTIONS = { ADD_TASK: 'add_task' };
/** The due choices of the "Add a task" form, in days from today. */
export const DUE_CHOICES = { none: null, today: 0, tomorrow: 1, week: 7 };
const MAX_EVENTS = 8;
const MAX_TASKS = 5;

/** The created devices the widget shows: the chosen one, or all of that kind. */
function chosen(monitor, setting, kind) {
  const created = monitor.createdDevices().filter((entry) => entry.kind === kind);
  const picked = created.find((entry) => entry.device.external_id === setting);
  return picked ? [picked] : created;
}

function addTaskButton(language, listExternalId) {
  // Form labels must be multi-language objects (the core picks the language).
  const label = textInEveryLanguage;
  return {
    type: 'button',
    label: pick(TEXT, language).addTask,
    icon: 'plus',
    style: 'secondary',
    action: {
      key: WIDGET_ACTIONS.ADD_TASK,
      params: { list: listExternalId },
      fields: [
        { key: 'title', type: 'string', required: true, label: label('taskTitle') },
        {
          key: 'due',
          type: 'select',
          default: 'none',
          label: label('due'),
          options: [
            { value: 'none', label: label('dueNone') },
            { value: 'today', label: label('dueToday') },
            { value: 'tomorrow', label: label('dueTomorrow') },
            { value: 'week', label: label('dueWeek') },
          ],
        },
      ],
    },
  };
}

/**
 * The widget content.
 * @param {{ monitor: ReturnType<import('./monitor.js').createMonitor>, gladys: object }} context
 * @param {{ settings?: object, language?: string, now: number }} request
 */
export function agendaContent({ monitor, gladys }, { settings = {}, language = 'en', now }) {
  const text = pick(TEXT, language);
  if (!monitor.signedIn) {
    return {
      ttl_seconds: 60,
      components: [{ type: 'text', variant: 'body', text: text.notSignedIn }],
    };
  }
  const calendars = chosen(monitor, settings.calendar, KINDS.CALENDAR);
  const lists = chosen(monitor, settings.list, KINDS.TODO_LIST);
  if (calendars.length === 0 && lists.length === 0) {
    return {
      ttl_seconds: 300,
      components: [{ type: 'text', variant: 'body', text: text.nothingToShow }],
    };
  }
  const components = [];

  if (calendars.length > 0) {
    const tomorrow = localMidnight(localDateKey(now), 1);
    const events = calendars
      .flatMap((entry) => monitor.cache.get(entry.device.external_id)?.events ?? [])
      .filter((event) => event.endMs > now && event.startMs < tomorrow)
      .sort((a, b) => a.startMs - b.startMs)
      .slice(0, MAX_EVENTS);
    const read = calendars.some((entry) => monitor.cache.has(entry.device.external_id));
    if (events.length > 0) {
      components.push({
        type: 'card-list',
        display: 'list',
        items: events.map((event) => ({
          title: truncate(event.title || '…', 60),
          subtitle: truncate(
            [formatRange(event, language), event.location].filter(Boolean).join(' · '),
            80,
          ),
          ...(event.startMs <= now
            ? { badge: { text: truncate(text.now, 16), color: WIDGET_COLORS.SUCCESS } }
            : {}),
        })),
      });
    } else {
      components.push({
        type: 'text',
        variant: 'caption',
        text: read ? text.noEventToday : text.syncing,
      });
    }
  }

  // Tasks: one list (the chosen one, or the only one added).
  const list = lists.length === 1 ? lists[0] : null;
  if (list) {
    const ids = gladys.externalIds(list.kind, list.platformId);
    const names = pick(NAMES, language);
    components.push(
      {
        type: 'value',
        label: names.open_tasks,
        icon: 'list',
        device_feature: ids.feature('open_tasks'),
      },
      {
        type: 'value',
        label: names.overdue_tasks,
        icon: 'alert-circle',
        device_feature: ids.feature('overdue_tasks'),
      },
    );
    const tasks = monitor.cache.get(list.device.external_id)?.tasks;
    if (tasks && tasks.length > 0) {
      const today = localDateKey(now);
      components.push({
        type: 'status',
        items: tasks.slice(0, MAX_TASKS).map((task) => ({
          label: truncate(task.title || '…', 40),
          value: truncate(
            task.dueDate ? formatDueDate(task.dueDate, language, now) : text.noDueDate,
            40,
          ),
          icon: task.dueDate && task.dueDate < today ? 'alert-circle' : 'circle',
          color:
            task.dueDate && task.dueDate < today
              ? WIDGET_COLORS.DANGER
              : task.dueDate === today
                ? WIDGET_COLORS.WARNING
                : WIDGET_COLORS.NEUTRAL,
        })),
      });
    } else if (tasks) {
      components.push({ type: 'text', variant: 'caption', text: text.noOpenTask });
    }
    components.push(addTaskButton(language, list.device.external_id));
  }
  return { ttl_seconds: 60, components };
}

/**
 * A tap on "Add a task" (form values validated by the core).
 * @returns {Promise<object>} the toast, in the configured language set
 */
export async function agendaAction(context, { actionKey, params = {}, values }) {
  if (actionKey !== WIDGET_ACTIONS.ADD_TASK) {
    throw new Error(`Unknown widget action: ${actionKey}`);
  }
  if (!values) {
    // A core without action forms (before 5.2.0) sends no values: nothing to add.
    throw new Error('This action needs Gladys 5.2.0 or later');
  }
  const due = Object.hasOwn(DUE_CHOICES, values.due) ? DUE_CHOICES[values.due] : null;
  const { title } = await addTask(context, {
    list: params.list,
    title: values.title,
    due_in_days: due,
  });
  return everyLanguage('taskAdded', truncate(title, 120));
}

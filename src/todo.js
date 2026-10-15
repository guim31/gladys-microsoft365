// -----------------------------------------------------------------------------
// Microsoft To Do lists: counts for the list device, the "task overdue"
// trigger, and finding the task a scene wants to complete.
//
// A To Do due date is a date, not an instant: a task due today is not overdue
// before tomorrow, whatever the hour. The counts follow the local calendar
// day, so they change at midnight without asking Graph.
// -----------------------------------------------------------------------------

import { fold } from './config.js';
import { graphDate } from './time.js';

/** One open Graph task, as the integration uses it; null when completed. */
export function normalizeTask(raw) {
  if (!raw || typeof raw !== 'object' || raw.status === 'completed' || !raw.id) {
    return null;
  }
  return {
    id: String(raw.id),
    title: String(raw.title ?? '').trim(),
    dueDate: raw.dueDateTime ? graphDate(raw.dueDateTime) : null,
    createdAt: Date.parse(raw.createdDateTime ?? '') || 0,
  };
}

/** The open tasks, the ones due first first, then the oldest. */
export function normalizeTasks(rawTasks) {
  return rawTasks
    .map(normalizeTask)
    .filter(Boolean)
    .sort(
      (a, b) =>
        (a.dueDate ?? '9999').localeCompare(b.dueDate ?? '9999') || a.createdAt - b.createdAt,
    );
}

/**
 * What the list device shows on `today` ('YYYY-MM-DD').
 * @returns {{ open: number, overdue: number, dueToday: number, overdueTasks: Array }}
 */
export function listState(tasks, today) {
  const overdueTasks = tasks.filter((task) => task.dueDate && task.dueDate < today);
  return {
    open: tasks.length,
    overdue: overdueTasks.length,
    dueToday: tasks.filter((task) => task.dueDate === today).length,
    overdueTasks,
  };
}

/**
 * The task a scene means by `text`: the one titled exactly so (case and
 * accents ignored; the one due first if there are several), else the only
 * one whose title contains it. Several partial matches: none — completing
 * the wrong task is worse than completing none.
 * @returns {{ task: object|null, reason: 'found'|'none'|'ambiguous' }}
 */
export function findTask(tasks, text) {
  const wanted = fold(text);
  if (!wanted) {
    return { task: null, reason: 'none' };
  }
  const exact = tasks.filter((task) => fold(task.title) === wanted);
  if (exact.length > 0) {
    return { task: exact[0], reason: 'found' };
  }
  const partial = tasks.filter((task) => fold(task.title).includes(wanted));
  if (partial.length === 1) {
    return { task: partial[0], reason: 'found' };
  }
  return { task: null, reason: partial.length > 1 ? 'ambiguous' : 'none' };
}

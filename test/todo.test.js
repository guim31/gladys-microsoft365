// -----------------------------------------------------------------------------
// To Do lists: due dates, counts, and finding the task a scene names.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { defaultTasks, fixture } from './helpers/world.js';
import { findTask, listState, normalizeTasks } from '../src/todo.js';

test('the tasks of the Graph reference are read', () => {
  const tasks = normalizeTasks(fixture('graph/todo-tasks.json').value);
  assert.deepEqual(
    tasks.map((task) => [task.title, task.dueDate]),
    [
      ['Shop for dinner', '2020-08-25'],
      ['Linked entity new task 1', null],
    ],
  );
});

test('completed tasks are left out; due first, then oldest', () => {
  const tasks = normalizeTasks(defaultTasks());
  assert.deepEqual(
    tasks.map((task) => [task.title, task.dueDate]),
    [
      ['Sortir les poubelles', '2026-10-09'],
      ['Arroser les plantes', '2026-10-11'],
      ['Acheter du pain', '2026-10-12'],
      ['Changer le filtre de la VMC', null],
    ],
  );
});

test('a task due today is not overdue before tomorrow', () => {
  const tasks = normalizeTasks(defaultTasks());
  const monday = listState(tasks, '2026-10-12');
  assert.equal(monday.open, 4);
  assert.equal(monday.overdue, 2);
  assert.equal(monday.dueToday, 1);
  assert.deepEqual(
    monday.overdueTasks.map((task) => task.id),
    ['task-bins', 'task-plants'],
  );
  const tuesday = listState(tasks, '2026-10-13');
  assert.equal(tuesday.overdue, 3);
  assert.equal(tuesday.dueToday, 0);
  assert.deepEqual(listState([], '2026-10-12'), {
    open: 0,
    overdue: 0,
    dueToday: 0,
    overdueTasks: [],
  });
});

test('finding a task: exact title first, else a single partial match', () => {
  const tasks = normalizeTasks([
    ...defaultTasks(),
    { id: 'task-bins-2', title: 'Sortir les poubelles jaunes', status: 'notStarted' },
  ]);
  assert.equal(findTask(tasks, 'sortir les POUBELLES').task.id, 'task-bins');
  assert.equal(findTask(tasks, 'pain').task.id, 'task-bread');
  assert.equal(findTask(tasks, 'filtre vmc').reason, 'none');
  // "poubelles" is in two titles, neither exactly: nothing is completed.
  assert.deepEqual(findTask(tasks, 'poubelles'), { task: null, reason: 'ambiguous' });
  assert.deepEqual(findTask(tasks, '  '), { task: null, reason: 'none' });
});

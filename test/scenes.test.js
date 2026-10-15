// -----------------------------------------------------------------------------
// Scene actions: add a task, complete a task, add an event — what reaches
// Microsoft Graph, and what the scene gets back.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createWorld } from './helpers/world.js';

const LIST = 'AAMkADIyAAAAABrJAAA=';

test('add a task with a due date: the local midnight, read back as that date', async () => {
  const world = await createWorld();
  try {
    const list = world.device('Tasks').external_id;
    const ack = await world.gladys.fake.sceneAction('add_task', {
      list,
      title: '  Racheter   du lait ',
      due_in_days: 1,
    });
    assert.deepEqual(ack, { success: true, data: { outputs: { title: 'Racheter du lait' } } });
    const created = world.microsoft.data.tasks[LIST].at(-1);
    assert.equal(created.title, 'Racheter du lait');
    // Tuesday 13 October, midnight in Paris.
    assert.deepEqual(created.dueDateTime, { dateTime: '2026-10-12T22:00:00', timeZone: 'UTC' });
    // The list is read again at once.
    assert.equal(world.state('Tasks', 'open_tasks'), 5);
  } finally {
    world.app.stop();
  }
});

test('add a task without a due date', async () => {
  const world = await createWorld();
  try {
    await world.gladys.fake.sceneAction('add_task', {
      list: world.device('Tasks').external_id,
      title: 'Appeler le plombier',
    });
    assert.equal(world.microsoft.data.tasks[LIST].at(-1).dueDateTime, undefined);
  } finally {
    world.app.stop();
  }
});

test('a wrong list, a calendar instead of a list, an empty title: a clear error', async () => {
  const world = await createWorld({ add: ['Tasks', 'Calendar'] });
  try {
    for (const fields of [
      { list: 'ext:test-integration:todo_list:unknown', title: 'x' },
      { list: world.device('Calendar').external_id, title: 'x' },
      { list: world.device('Monthly Tasks').external_id, title: 'x' }, // not added
    ]) {
      const ack = await world.gladys.fake.sceneAction('add_task', fields);
      assert.equal(ack.success, false);
      assert.match(ack.error, /liste To Do ajoutée à Gladys/);
    }
    const ack = await world.gladys.fake.sceneAction('add_task', {
      list: world.device('Tasks').external_id,
      title: '   ',
    });
    assert.equal(ack.success, false);
    assert.match(ack.error, /titre est vide/);
  } finally {
    world.app.stop();
  }
});

test('complete a task by its title; nothing when unsure', async () => {
  const world = await createWorld();
  try {
    const list = world.device('Tasks').external_id;
    let ack = await world.gladys.fake.sceneAction('complete_task', {
      list,
      title: 'acheter du PAIN',
    });
    assert.deepEqual(ack.data.outputs, { completed: 1, title: 'Acheter du pain' });
    const bread = world.microsoft.data.tasks[LIST].find((task) => task.id === 'task-bread');
    assert.equal(bread.status, 'completed');
    assert.equal(world.state('Tasks', 'due_today_tasks'), 0);

    ack = await world.gladys.fake.sceneAction('complete_task', { list, title: 'licorne' });
    assert.deepEqual(ack.data.outputs, { completed: 0, title: '' });
    // "les" is in two titles: no guess.
    ack = await world.gladys.fake.sceneAction('complete_task', { list, title: 'les' });
    assert.deepEqual(ack.data.outputs, { completed: 0, title: '' });
  } finally {
    world.app.stop();
  }
});

test('add an event: whole minutes, busy, the calendar read again', async () => {
  const world = await createWorld();
  try {
    world.clock.now += 20_000; // 08:00:20
    const ack = await world.gladys.fake.sceneAction('add_event', {
      calendar: world.device('Famille').external_id,
      title: 'Ménage',
      start_in_minutes: 0,
      duration_minutes: 90,
    });
    assert.deepEqual(ack.data.outputs, { start: '08:01' });
    const [event] = world.microsoft.data.events['AAMkAGI2TGuLAAB=family'];
    assert.equal(event.subject, 'Ménage');
    assert.equal(event.showAs, 'busy');
    assert.deepEqual(event.start, { dateTime: '2026-10-12T06:01:00', timeZone: 'UTC' });
    assert.deepEqual(event.end, { dateTime: '2026-10-12T07:31:00', timeZone: 'UTC' });
    assert.equal(world.state('Famille', 'next_event'), 'Ménage · 08:01');
  } finally {
    world.app.stop();
  }
});

test('not signed in: the scene action fails with the reason', async () => {
  const world = await createWorld();
  try {
    const list = world.device('Tasks').external_id;
    await world.gladys.fake.action('sign_out', {});
    const ack = await world.gladys.fake.sceneAction('add_task', { list, title: 'x' });
    assert.equal(ack.success, false);
    assert.match(ack.error, /Pas connecté/);
  } finally {
    world.app.stop();
  }
});

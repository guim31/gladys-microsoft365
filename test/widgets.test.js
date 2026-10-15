// -----------------------------------------------------------------------------
// Dashboard widget "Agenda and tasks": the content fits the core vocabulary
// and budget, its live tiles point at real features, and its button adds a
// task.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { validateWidgetContent } from '@gladysassistant/integration-sdk';
import { AGENDA_WIDGET } from '../src/widgets.js';
import { createWorld } from './helpers/world.js';

async function content(world, settings = {}, language = 'fr') {
  const ack = await world.gladys.fake.widgetGet(AGENDA_WIDGET, {
    settings,
    language,
    units: 'metric',
  });
  assert.equal(ack.success, true, ack.error);
  const result = ack.data.content ?? ack.data;
  assert.deepEqual(validateWidgetContent(result), [], 'rendered exactly as sent');
  return result;
}

const ofType = (result, type) => result.components.filter((component) => component.type === type);

test("today's agenda and the tasks of the list, in the dashboard language", async () => {
  const world = await createWorld({ add: ['Calendar', 'Tasks'] });
  try {
    const result = await content(world);
    const [agenda] = ofType(result, 'card-list');
    assert.deepEqual(
      agenda.items.map((item) => [item.title, item.subtitle, item.badge?.text]),
      [
        ['Télétravail', 'toute la journée', 'En cours'],
        ["Point d'équipe", '09:30 – 10:00 · Salle Rivoli', undefined],
        ['Dentiste', '14:00 – 14:45 · Cabinet du Dr Durand', undefined],
      ],
    );
    // Live tiles bound to features that exist on the list device.
    const features = new Set(world.device('Tasks').features.map((feature) => feature.external_id));
    const tiles = ofType(result, 'value');
    assert.equal(tiles.length, 2);
    for (const tile of tiles) {
      assert.ok(features.has(tile.device_feature), tile.device_feature);
    }
    const [status] = ofType(result, 'status');
    assert.deepEqual(
      status.items.map((item) => [item.label, item.value, item.color]),
      [
        ['Sortir les poubelles', 'ven. 9 oct.', 'danger'],
        ['Arroser les plantes', 'dim. 11 oct.', 'danger'],
        ['Acheter du pain', "aujourd'hui", 'warning'],
        ['Changer le filtre de la VMC', 'sans échéance', 'neutral'],
      ],
    );
    const [button] = ofType(result, 'button');
    assert.equal(button.action.key, 'add_task');
    assert.deepEqual(button.action.params, { list: world.device('Tasks').external_id });
    assert.deepEqual(
      button.action.fields.map((field) => field.key),
      ['title', 'due'],
    );

    const english = await content(world, {}, 'en');
    assert.equal(ofType(english, 'card-list')[0].items[0].subtitle, 'all day');
  } finally {
    world.app.stop();
  }
});

test('the settings pick one calendar and one list among several', async () => {
  const world = await createWorld();
  try {
    // Two lists added and none chosen: no task part.
    let result = await content(world);
    assert.equal(ofType(result, 'button').length, 0);
    result = await content(world, {
      calendar: world.device('Famille').external_id,
      list: world.device('Monthly Tasks').external_id,
    });
    assert.deepEqual(
      ofType(result, 'text').map((component) => component.text),
      ["Plus rien à l'agenda aujourd'hui.", 'Aucune tâche ouverte.'],
    );
    assert.equal(ofType(result, 'button').length, 1);
  } finally {
    world.app.stop();
  }
});

test('the agenda keeps what is left of today, 8 events at most', async () => {
  const many = Array.from({ length: 12 }, (_, index) => ({
    id: `event-${index}`,
    subject: `Créneau ${index}`,
    isAllDay: false,
    showAs: 'busy',
    start: { dateTime: `2026-10-12T${String(8 + index).padStart(2, '0')}:00:00`, timeZone: 'UTC' },
    end: { dateTime: `2026-10-12T${String(8 + index).padStart(2, '0')}:30:00`, timeZone: 'UTC' },
  }));
  const world = await createWorld({
    add: ['Calendar'],
    microsoft: { events: { 'AAMkAGI2TGuLAAA=': many } },
  });
  try {
    const [agenda] = ofType(await content(world), 'card-list');
    assert.equal(agenda.items.length, 8);
    assert.equal(agenda.items[0].title, 'Créneau 0');
  } finally {
    world.app.stop();
  }
});

test('nothing added, not signed in: one sentence that says what to do', async () => {
  let world = await createWorld({ add: 'none' });
  try {
    const result = await content(world);
    assert.match(result.components[0].text, /onglet « Découverte »/);
  } finally {
    world.app.stop();
  }
  world = await createWorld({ signedIn: false });
  try {
    const result = await content(world, {}, 'de');
    assert.match(result.components[0].text, /Nicht bei Microsoft angemeldet/);
  } finally {
    world.app.stop();
  }
});

test('the "Add a task" button adds the task and answers with a toast', async () => {
  const world = await createWorld();
  try {
    const list = world.device('Tasks').external_id;
    const ack = await world.gladys.fake.widgetAction(
      AGENDA_WIDGET,
      'add_task',
      { list },
      { settings: {}, values: { title: 'Payer la cantine', due: 'week' } },
    );
    assert.equal(ack.success, true);
    assert.equal((ack.data.message ?? ack.data).fr, 'Tâche « Payer la cantine » ajoutée.');
    const created = world.microsoft.data.tasks['AAMkADIyAAAAABrJAAA='].at(-1);
    assert.deepEqual(created.dueDateTime, { dateTime: '2026-10-18T22:00:00', timeZone: 'UTC' });

    // An older core sends no form values: refused rather than an empty task.
    const old = await world.gladys.fake.widgetAction(AGENDA_WIDGET, 'add_task', { list });
    assert.equal(old.success, false);
  } finally {
    world.app.stop();
  }
});

// -----------------------------------------------------------------------------
// Consistency checks between `gladys-assistant-integration.json` and the code.
// The manifest is validated by the store indexer, but nothing there can know
// which handlers the code actually registers — these tests keep both in sync.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { ACTIONS } from '../src/app.js';
import { UPCOMING_MINUTES } from '../src/calendar.js';
import { DEFAULT_CONFIG, LANGUAGES, REFRESH_MINUTES } from '../src/config.js';
import {
  SCENE_ACTIONS,
  SCENE_TRIGGER_KEYS,
  eventTriggerData,
  overdueTriggerData,
} from '../src/scenes.js';
import { AGENDA_WIDGET } from '../src/widgets.js';

const manifest = JSON.parse(
  await readFile(new URL('../gladys-assistant-integration.json', import.meta.url), 'utf8'),
);
const pkg = JSON.parse(await readFile(new URL('../package.json', import.meta.url), 'utf8'));

const keysOf = (list) => (list ?? []).map((entry) => entry.key);
const field = (list, key) => list.find((entry) => entry.key === key);

// Every list of form fields the manifest declares (same field grammar).
const allFields = [
  ...manifest.config_schema,
  ...[...manifest.actions, ...manifest.scene_triggers, ...manifest.scene_actions].flatMap(
    (item) => item.fields ?? [],
  ),
  ...manifest.widgets.flatMap((widget) => widget.settings ?? []),
];

function minGladysVersion() {
  const match = manifest.gladys_version.match(/>=\s*(\d+)\.(\d+)\.(\d+)/);
  assert.ok(match, 'gladys_version must declare a minimum version');
  return match.slice(1).map(Number);
}

test('a device integration of the services shelf, cloud only, for Gladys 5.2.0+', () => {
  // The widget's "Add a task" button opens a form: Gladys 5.2.0 (#3168).
  assert.equal(manifest.type, 'device');
  assert.equal(manifest.name, 'Microsoft 365');
  assert.deepEqual(minGladysVersion(), [5, 2, 0]);
  assert.deepEqual(manifest.categories, ['services']);
  assert.deepEqual(manifest.transports, ['cloud']);
  assert.equal(
    manifest.docker_image,
    `ghcr.io/guim31/gladys-microsoft365:${manifest.version}`,
    'the image of this repository, tagged with the version',
  );
  assert.equal(
    manifest.cover_image,
    'https://raw.githubusercontent.com/guim31/gladys-microsoft365/main/cover.png',
  );
});

test('the manifest version is the package version', () => {
  // The Release workflow writes both: a mismatch means one was edited by hand.
  assert.equal(manifest.version, pkg.version);
});

test('the catalog description holds 10 to 100 characters per language', () => {
  for (const [lang, text] of Object.entries(manifest.description)) {
    assert.ok(
      text.length >= 10 && text.length <= 100,
      `description.${lang} has ${text.length} characters`,
    );
  }
});

test('every manifest action has a handler, and vice versa', () => {
  assert.deepEqual(keysOf(manifest.actions).sort(), Object.keys(ACTIONS).sort());
  for (const action of manifest.actions) {
    // An action applies no default and cannot fill a secret (Gladys#3154, #3155).
    assert.equal(action.fields, undefined, `${action.key} needs no field`);
  }
});

test('every scene action has a handler, and vice versa', () => {
  assert.deepEqual(keysOf(manifest.scene_actions).sort(), Object.keys(SCENE_ACTIONS).sort());
});

test('every scene trigger the code fires is declared, and vice versa', () => {
  assert.deepEqual(keysOf(manifest.scene_triggers).sort(), [...SCENE_TRIGGER_KEYS].sort());
});

test('the data of every scene event matches the declared filters and variables', () => {
  const event = {
    title: 'Dentiste',
    location: 'Cabinet',
    startMs: Date.parse('2026-10-12T12:00:00Z'),
    endMs: Date.parse('2026-10-12T12:45:00Z'),
    allDay: false,
  };
  const device = { external_id: 'ext:x:calendar:1', name: 'Calendar (Outlook)' };
  const samples = {
    event_started: eventTriggerData({ key: 'event_started', event }, device, ['dentiste'], 'fr'),
    event_ended: eventTriggerData({ key: 'event_ended', event }, device, [], 'fr'),
    event_upcoming: eventTriggerData(
      { key: 'event_upcoming', event, minutes: 15 },
      device,
      [],
      'fr',
    ),
    task_overdue: overdueTriggerData(
      { title: 'Pain', dueDate: '2026-10-11' },
      { external_id: 'ext:x:todo_list:1', name: 'Tasks (To Do)' },
      'fr',
      Date.parse('2026-10-12T06:00:00Z'),
    ),
  };
  for (const trigger of manifest.scene_triggers) {
    const data = samples[trigger.key];
    const declared = new Set([...keysOf(trigger.fields), ...keysOf(trigger.variables)]);
    assert.deepEqual(
      Object.keys(data).sort(),
      [...declared].sort(),
      `${trigger.key}: the data carries exactly the declared keys`,
    );
    assert.ok(Object.keys(data).length <= 30);
    for (const [key, value] of Object.entries(data)) {
      assert.ok(
        value === null || ['string', 'number', 'boolean'].includes(typeof value),
        `${trigger.key}.${key} is flat`,
      );
    }
  }
});

test('the "in N minutes" choices are the lead times the code computes', () => {
  const minutes = field(
    manifest.scene_triggers.find((trigger) => trigger.key === 'event_upcoming').fields,
    'minutes',
  );
  assert.equal(minutes.required, true, 'otherwise a scene would run once per lead time');
  assert.deepEqual(
    minutes.options.map((option) => Number(option.value)),
    UPCOMING_MINUTES,
  );
});

test('the widget is declared and handled', () => {
  assert.deepEqual(keysOf(manifest.widgets), [AGENDA_WIDGET]);
  for (const widget of manifest.widgets) {
    assert.ok(widget.label.en.length <= 30);
    assert.ok(widget.description.en.length <= 100);
  }
});

test('config_schema defaults and choices are the ones of the code', () => {
  for (const entry of manifest.config_schema.filter((item) => item.type !== 'section')) {
    if (entry.default !== undefined) {
      assert.equal(DEFAULT_CONFIG[entry.key], entry.default, `default of ${entry.key}`);
    }
    assert.ok(entry.key in DEFAULT_CONFIG, `${entry.key} is unknown to the code`);
  }
  const options = (key) => field(manifest.config_schema, key).options.map((option) => option.value);
  assert.deepEqual(options('refresh_minutes'), REFRESH_MINUTES);
  assert.deepEqual(options('language'), LANGUAGES);
  assert.equal(DEFAULT_CONFIG.language, 'fr');
  // The client ID of a public application is no secret; no secret field at all.
  assert.ok(allFields.every((entry) => entry.type !== 'secret'));
});

test('section fields are purely presentational, with https links', () => {
  for (const section of allFields.filter((entry) => entry.type === 'section')) {
    assert.equal(section.required, undefined);
    assert.equal(section.default, undefined);
    assert.equal(section.placeholder, undefined);
    for (const language of LANGUAGES) {
      assert.ok(section.description[language].length <= 1000);
    }
    for (const link of section.links ?? []) {
      assert.match(link.url, /^https:\/\//);
    }
  }
});

test('dynamic selects use the devices source, never with static options', () => {
  const dynamic = allFields.filter((entry) => entry.source !== undefined);
  assert.ok(dynamic.length >= 5);
  for (const entry of dynamic) {
    assert.equal(entry.source, 'devices');
    assert.equal(entry.options, undefined);
    assert.equal(entry.default, undefined);
  }
});

test('number fields have whole bounds and defaults (no step in the Gladys form)', () => {
  for (const entry of allFields.filter((item) => item.type === 'number')) {
    for (const key of ['min', 'max', 'default']) {
      if (entry[key] !== undefined) {
        assert.ok(Number.isInteger(entry[key]), `${entry.key}.${key}`);
      }
    }
  }
});

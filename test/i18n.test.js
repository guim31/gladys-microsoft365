// -----------------------------------------------------------------------------
// Every text Gladys shows exists in the four languages of its interface
// (English, French, German, Spanish): the manifest, the message tables of
// the code and the documentation. German says "du", never the polite "Sie".
// -----------------------------------------------------------------------------

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { LANGUAGES } from '../src/config.js';
import { MESSAGES, NAMES, TEXT } from '../src/i18n.js';

const manifest = JSON.parse(
  await readFile(new URL('../gladys-assistant-integration.json', import.meta.url), 'utf8'),
);

const LANGS = ['en', 'fr', 'de', 'es'];
const POLITE_GERMAN = /(?<!\p{L})(Sie|Ihr|Ihre|Ihren|Ihrem|Ihrer|Ihres)(?!\p{L})/u;

// Upper bounds, in characters: what the Gladys interface shows without cutting.
const MAX = {
  catalog: 100,
  widgetLabel: 30,
  label: 60,
  placeholder: 60,
  description: 300,
  section: 1000,
  message: 200,
  short: 24,
  featureName: 40,
};

const sorted = (keys) => [...keys].sort();

/** Checks one multi-language object; returns its problems. */
function problems(value, path, max) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return [`${path} must be a multi-language object`];
  }
  const found = [];
  if (JSON.stringify(sorted(Object.keys(value))) !== JSON.stringify(sorted(LANGS))) {
    found.push(`${path} has [${Object.keys(value)}], expected exactly [${LANGS}]`);
  }
  for (const [lang, text] of Object.entries(value)) {
    found.push(...textProblems(text, `${path}.${lang}`, lang, max));
  }
  return found;
}

function textProblems(text, path, lang, max) {
  if (typeof text !== 'string' || text.trim() === '') {
    return [`${path} is empty`];
  }
  const found = [];
  if (text.length > max) {
    found.push(`${path} has ${text.length} characters (${max} at most)`);
  }
  if (lang === 'de' && POLITE_GERMAN.test(text)) {
    found.push(`${path} uses the polite form: "${text}"`);
  }
  return found;
}

/** Every multi-language object of the manifest, with its length limit. */
function manifestTexts(root) {
  const texts = [];
  const walk = (value, path, key, parent) => {
    if (Array.isArray(value)) {
      value.forEach((item, index) => walk(item, `${path}[${index}]`, key, parent));
      return;
    }
    if (!value || typeof value !== 'object') {
      return;
    }
    const isText =
      ['label', 'description', 'placeholder'].includes(key) ||
      Object.keys(value).some((child) => LANGS.includes(child));
    if (isText) {
      let max = MAX[key] ?? MAX.label;
      if (path === 'manifest.description' || /^manifest\.widgets\[\d+\]\.description$/.test(path)) {
        max = MAX.catalog;
      } else if (/^manifest\.widgets\[\d+\]\.label$/.test(path)) {
        max = MAX.widgetLabel;
      } else if (key === 'description' && parent?.type === 'section') {
        max = MAX.section;
      }
      texts.push({ path, value, max });
      return;
    }
    for (const [child, item] of Object.entries(value)) {
      walk(item, `${path}.${child}`, child, value);
    }
  };
  walk(root, 'manifest', null, null);
  return texts;
}

// Tables written one language at a time: same keys everywhere, no empty text.
function tableProblems(table, name, max, render = (value) => value) {
  const found = [];
  if (JSON.stringify(sorted(Object.keys(table))) !== JSON.stringify(sorted(LANGS))) {
    found.push(`${name} has [${Object.keys(table)}], expected exactly [${LANGS}]`);
  }
  for (const lang of LANGS) {
    const entries = table[lang] ?? {};
    if (
      JSON.stringify(sorted(Object.keys(entries))) !== JSON.stringify(sorted(Object.keys(table.en)))
    ) {
      found.push(`${name}.${lang} does not have the keys of ${name}.en`);
    }
    for (const [key, value] of Object.entries(entries)) {
      found.push(...textProblems(render(value), `${name}.${lang}.${key}`, lang, max(key)));
    }
  }
  return found;
}

test('the configuration offers the four languages', () => {
  assert.deepEqual(LANGUAGES, LANGS);
});

test('every text of the manifest is in English, French, German and Spanish', () => {
  const texts = manifestTexts(manifest);
  assert.ok(texts.length >= 100, `only ${texts.length} texts found in the manifest`);
  const found = texts.flatMap(({ path, value, max }) => problems(value, path, max));
  assert.deepEqual(found, []);
});

test('the name is the same in every language, 3 to 30 characters', () => {
  assert.equal(typeof manifest.name, 'string');
  assert.ok(manifest.name.length >= 3 && manifest.name.length <= 30);
});

test('feature names, short texts and messages exist in every language', () => {
  assert.deepEqual(
    tableProblems(NAMES, 'NAMES', () => MAX.featureName),
    [],
  );
  const long = new Set(['notSignedIn', 'nothingToShow', 'noEventToday', 'syncing']);
  assert.deepEqual(
    tableProblems(TEXT, 'TEXT', (key) => (long.has(key) ? MAX.message : MAX.short)),
    [],
  );
  // Messages rendered with long but realistic values.
  const render = (value) =>
    typeof value === 'function'
      ? value('https://login.microsoft.com/device', 'ABCDEFGHJ', 15)
      : value;
  assert.deepEqual(
    tableProblems(MESSAGES, 'MESSAGES', () => MAX.message, render),
    [],
  );
  const report = (lang) => MESSAGES[lang].report('Camille Martin', 3, 2);
  for (const lang of LANGS) {
    assert.ok(report(lang).length <= MAX.message);
  }
});

test('the guard catches a missing language, an empty text and the polite German form', () => {
  assert.equal(problems({ en: 'a', fr: 'a', de: 'a' }, 'x', 60).length, 1);
  assert.equal(problems({ en: 'a', fr: 'a', de: '', es: 'a' }, 'x', 60).length, 1);
  assert.equal(problems({ en: 'a', fr: 'a', de: 'a', es: 'a', it: 'a' }, 'x', 60).length, 1);
  assert.equal(problems({ en: 'a', fr: 'a', de: 'a'.repeat(61), es: 'a' }, 'x', 60).length, 1);
  for (const polite of ['Wählen Sie', 'Ihr Konto', 'Ihre Liste', 'mit Ihrem', 'Ihren', 'Ihrer']) {
    assert.equal(problems({ en: 'a', fr: 'a', de: polite, es: 'a' }, 'x', 60).length, 1, polite);
  }
  assert.deepEqual(
    problems({ en: 'a', fr: 'a', de: 'Siebenschläfer, ihre', es: 'a' }, 'x', 60),
    [],
  );
  const forgotten = structuredClone(manifest);
  delete forgotten.scene_actions[0].fields[1].label.de;
  const found = manifestTexts(forgotten).flatMap(({ path, value, max }) =>
    problems(value, path, max),
  );
  assert.deepEqual(found, [
    'manifest.scene_actions[0].fields[1].label has [en,fr,es], expected exactly [en,fr,de,es]',
  ]);
});

// --- Documentation -----------------------------------------------------------

const docs = Object.fromEntries(
  await Promise.all(
    LANGS.map(async (lang) => [
      lang,
      await readFile(new URL(`../docs/${lang}.md`, import.meta.url), 'utf8'),
    ]),
  ),
);
const headings = (text, level) =>
  text.split('\n').filter((line) => line.startsWith(`${'#'.repeat(level)} `));

test('the four docs have the same sections, in the same order', () => {
  for (const level of [2, 3]) {
    const counts = LANGS.map((lang) => headings(docs[lang], level).length);
    assert.ok(
      counts.every((count) => count === counts[0]),
      `level ${level} headings: ${LANGS.map((lang, i) => `${lang} ${counts[i]}`).join(', ')}`,
    );
  }
  assert.ok(headings(docs.en, 2).length >= 8, 'the sections the brief asks for');
});

test('the docs name every action, trigger, scene action, widget and feature', () => {
  const keys = [
    ...manifest.actions,
    ...manifest.scene_triggers,
    ...manifest.scene_actions,
    ...manifest.widgets,
  ];
  for (const lang of LANGS) {
    for (const entry of keys) {
      assert.ok(docs[lang].includes(entry.label[lang]), `docs/${lang}.md: "${entry.label[lang]}"`);
    }
    for (const name of Object.values(NAMES[lang])) {
      assert.ok(docs[lang].includes(name), `docs/${lang}.md: feature "${name}"`);
    }
    assert.ok(docs[lang].length >= 300);
    assert.match(docs[lang], /Microsoft/);
  }
});

test('the docs say it was developed without a test account, and is not affiliated', () => {
  assert.match(docs.en, /Developed without a test account: feedback welcome/);
  assert.match(docs.en, /not affiliated with Microsoft/i);
  assert.match(docs.fr, /sans lien avec Microsoft/i);
  assert.match(docs.de, /nicht mit Microsoft verbunden/i);
  assert.match(docs.es, /sin relación con Microsoft/i);
  for (const lang of ['de']) {
    for (const line of docs[lang].split('\n')) {
      assert.doesNotMatch(line, POLITE_GERMAN, `docs/de.md: ${line}`);
    }
  }
});

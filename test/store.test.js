// -----------------------------------------------------------------------------
// /data: the tokens readable by the container user only, the state kept
// across restarts, a broken file never fatal.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createStore } from '../src/store.js';

process.env.LOG_LEVEL ??= 'silent';

const TOKENS = {
  client_id: '0f1e2d3c-4b5a-4968-8776-655443322110',
  tenant: 'common',
  refresh_token: 'invented-refresh-token',
  account: { id: 'account', name: 'Camille' },
};

test('tokens: written with mode 600, read back, deleted on sign-out', async () => {
  const directory = mkdtempSync(join(tmpdir(), 'ms365-'));
  const store = createStore(directory);
  await store.load();
  assert.equal(store.getTokens(), null);
  await store.setTokens(TOKENS);
  const file = join(directory, 'tokens.json');
  assert.equal(statSync(file).mode & 0o777, 0o600);
  assert.deepEqual(JSON.parse(readFileSync(file, 'utf8')), TOKENS);

  const again = createStore(directory);
  await again.load();
  assert.deepEqual(again.getTokens(), TOKENS);
  await again.setTokens(null);
  const third = createStore(directory);
  await third.load();
  assert.equal(third.getTokens(), null);
});

test('state: Graph ids and announced overdue tasks survive a restart', async () => {
  const directory = mkdtempSync(join(tmpdir(), 'ms365-'));
  const store = createStore(directory);
  await store.load();
  store.setGraphIds({ abc: 'AAMkAGI2TGuLAAA=' });
  store.setOverdue('list', ['task-1']);
  await store.save();
  assert.equal(statSync(join(directory, 'state.json')).mode & 0o777, 0o600);
  const again = createStore(directory);
  await again.load();
  assert.equal(again.graphId('abc'), 'AAMkAGI2TGuLAAA=');
  assert.deepEqual(again.overdue('list'), ['task-1']);
  assert.equal(again.overdue('other'), undefined);
});

test('a broken file starts afresh; an unwritable /data is not fatal', async () => {
  const directory = mkdtempSync(join(tmpdir(), 'ms365-'));
  writeFileSync(join(directory, 'state.json'), '{not json');
  writeFileSync(join(directory, 'tokens.json'), '[]');
  const store = createStore(directory);
  await store.load();
  assert.equal(store.getTokens(), null);
  assert.equal(store.graphId('x'), undefined);

  // A "directory" under a regular file: every write fails (ENOTDIR).
  const readOnly = createStore(join(directory, 'state.json', 'data'));
  await readOnly.load();
  await readOnly.setTokens(TOKENS);
  await readOnly.save();
  assert.deepEqual(readOnly.getTokens(), TOKENS, 'kept in memory');
});

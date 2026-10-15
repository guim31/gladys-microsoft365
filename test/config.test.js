// -----------------------------------------------------------------------------
// Configuration: every value checked before the code sees it.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_CONFIG, normalizeConfig, parseKeywords, refreshMs } from '../src/config.js';

test('defaults', () => {
  assert.deepEqual(normalizeConfig(), {
    client_id: '',
    tenant: 'common',
    refresh_minutes: '5',
    trigger_keywords: [],
    language: 'fr',
  });
  assert.equal(refreshMs(normalizeConfig()), 300_000);
  assert.equal(DEFAULT_CONFIG.language, 'fr');
});

test('the client ID must be a GUID; the tenant cannot leave the URL path', () => {
  const config = normalizeConfig({
    client_id: ' 0F1E2D3C-4B5A-4968-8776-655443322110 ',
    tenant: 'Contoso.onmicrosoft.com',
  });
  assert.equal(config.client_id, '0f1e2d3c-4b5a-4968-8776-655443322110');
  assert.equal(config.tenant, 'contoso.onmicrosoft.com');
  assert.equal(normalizeConfig({ client_id: 'my app' }).client_id, '');
  assert.equal(normalizeConfig({ tenant: '../../evil?x=' }).tenant, 'common');
  assert.equal(normalizeConfig({ tenant: '' }).tenant, 'common');
});

test('unknown choices fall back to the defaults', () => {
  const config = normalizeConfig({ refresh_minutes: '3', language: 'it' });
  assert.equal(config.refresh_minutes, '5');
  assert.equal(config.language, 'fr');
  assert.equal(normalizeConfig({ refresh_minutes: 1 }).refresh_minutes, '1');
});

test('keywords: trimmed, deduplicated (case and accents ignored), in order', () => {
  assert.deepEqual(parseKeywords(' Télétravail, dentiste;TELETRAVAIL\n Congés ,, '), [
    'Télétravail',
    'dentiste',
    'Congés',
  ]);
  assert.deepEqual(parseKeywords(undefined), []);
});

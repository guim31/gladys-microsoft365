// -----------------------------------------------------------------------------
// Integration configuration.
//
// Filled in by the administrator from the `config_schema` of the manifest. The
// defaults below MUST stay the `default` values declared there
// (test/manifest.test.js checks it). Nothing secret lives here: the client ID
// of a public application is not a secret, and the Microsoft tokens are kept
// in /data (src/store.js), never in the configuration.
// -----------------------------------------------------------------------------

export const LANGUAGES = ['en', 'fr', 'de', 'es'];

/** Minutes between two reads of Microsoft Graph (the `refresh_minutes` select). */
export const REFRESH_MINUTES = ['1', '2', '5', '10', '15'];

export const DEFAULT_CONFIG = {
  client_id: '',
  tenant: 'common',
  refresh_minutes: '5',
  trigger_keywords: '',
  language: 'fr',
};

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
// A tenant alias (common, consumers, organizations), a directory GUID or a
// domain (contoso.onmicrosoft.com): it goes into a URL path, nothing else.
const TENANT = /^[a-z0-9][a-z0-9.-]{0,99}$/i;

/** True when `value` looks like an application (client) ID. */
export function isClientId(value) {
  return GUID.test(String(value ?? '').trim());
}

/**
 * The keywords watched by the scene triggers, as typed (trimmed), in order.
 * @param {unknown} value comma, semicolon or line separated
 */
export function parseKeywords(value) {
  const seen = new Set();
  const keywords = [];
  for (const raw of String(value ?? '').split(/[,;\n]/)) {
    const keyword = raw.trim().slice(0, 60);
    if (keyword && !seen.has(fold(keyword))) {
      seen.add(fold(keyword));
      keywords.push(keyword);
    }
  }
  return keywords.slice(0, 20);
}

/** Lower case, without accents nor repeated spaces: for comparisons only. */
export function fold(text) {
  return String(text ?? '')
    .normalize('NFD')
    .replace(/\p{M}/gu, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Merge the user config with the defaults, with every value checked: the rest
 * of the code never sees an unknown language or a tenant with a slash.
 * @param {Record<string, unknown>} [raw] config returned by the SDK
 */
export function normalizeConfig(raw = {}) {
  const value = (key) => String(raw[key] ?? DEFAULT_CONFIG[key]).trim();
  const tenant = value('tenant');
  return {
    client_id: isClientId(raw.client_id) ? String(raw.client_id).trim().toLowerCase() : '',
    tenant: TENANT.test(tenant) ? tenant.toLowerCase() : DEFAULT_CONFIG.tenant,
    refresh_minutes: REFRESH_MINUTES.includes(value('refresh_minutes'))
      ? value('refresh_minutes')
      : DEFAULT_CONFIG.refresh_minutes,
    trigger_keywords: parseKeywords(raw.trigger_keywords),
    language: LANGUAGES.includes(value('language')) ? value('language') : DEFAULT_CONFIG.language,
  };
}

/** Milliseconds between two reads of Microsoft Graph. */
export function refreshMs(config) {
  return Number(config.refresh_minutes) * 60_000;
}

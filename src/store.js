// -----------------------------------------------------------------------------
// What the integration keeps across restarts, in /data (the only writable
// place of the sandbox):
//   - tokens.json (mode 600): the Microsoft refresh token, the client ID and
//     tenant it was issued for, and the account it belongs to (id and display
//     name). The access token is never written: it lives one hour, in memory.
//   - state.json: the Graph id behind each published device, and the open
//     tasks already announced as overdue (so a restart neither forgets nor
//     repeats a "task overdue" trigger).
// DATA_DIR overrides /data (tests, local runs).
// -----------------------------------------------------------------------------

import { chmod, mkdir, readFile, rename, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { createLogger } from '@gladysassistant/integration-sdk';

const logger = createLogger({ name: 'store' });

async function readJson(file) {
  try {
    const parsed = JSON.parse(await readFile(file, 'utf8'));
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : null;
  } catch (err) {
    if (err.code !== 'ENOENT') {
      logger.warn(`Could not read ${file}, starting afresh (${err.code ?? err.name})`);
    }
    return null;
  }
}

async function writeJson(directory, file, value, mode) {
  await mkdir(directory, { recursive: true });
  const temporary = `${file}.tmp`;
  await writeFile(temporary, JSON.stringify(value), { mode });
  await chmod(temporary, mode);
  await rename(temporary, file);
}

/**
 * @param {string} [directory] defaults to DATA_DIR or /data
 */
export function createStore(directory = process.env.DATA_DIR || '/data') {
  const stateFile = join(directory, 'state.json');
  const tokenFile = join(directory, 'tokens.json');
  let state = { ids: {}, overdue: {} };
  let tokens = null;
  let warned = false;

  const warnOnce = (file, err) => {
    if (!warned) {
      warned = true;
      logger.warn(`Could not write ${file} (${err.code ?? err.name})`);
    }
  };

  return {
    async load() {
      const loaded = await readJson(stateFile);
      state = {
        ids: loaded?.ids && typeof loaded.ids === 'object' ? loaded.ids : {},
        overdue: loaded?.overdue && typeof loaded.overdue === 'object' ? loaded.overdue : {},
      };
      const loadedTokens = await readJson(tokenFile);
      tokens = typeof loadedTokens?.refresh_token === 'string' ? loadedTokens : null;
    },

    /** @returns {{ client_id, tenant, refresh_token, account: { id, name } } | null} */
    getTokens() {
      return tokens;
    },

    /** Replaces the tokens (null forgets them, as a sign-out does). */
    async setTokens(value) {
      tokens = value;
      try {
        if (value === null) {
          await rm(tokenFile, { force: true });
        } else {
          await writeJson(directory, tokenFile, value, 0o600);
        }
      } catch (err) {
        // Not fatal: the session keeps working until the next restart.
        warnOnce(tokenFile, err);
      }
    },

    /** The Graph id behind a device platform id (from the last discovery). */
    graphId(platformId) {
      return state.ids[platformId];
    },

    setGraphIds(entries) {
      state.ids = { ...state.ids, ...entries };
    },

    /** The task ids already announced as overdue for a list, or undefined. */
    overdue(platformId) {
      return state.overdue[platformId];
    },

    setOverdue(platformId, taskIds) {
      state.overdue[platformId] = taskIds;
    },

    async save() {
      try {
        await writeJson(directory, stateFile, state, 0o600);
      } catch (err) {
        warnOnce(stateFile, err);
      }
    },
  };
}

/** A store kept in memory (tests). */
export function createMemoryStore({ tokens = null } = {}) {
  let current = tokens;
  const state = { ids: {}, overdue: {} };
  return {
    saves: 0,
    async load() {},
    getTokens: () => current,
    async setTokens(value) {
      current = value;
    },
    graphId: (platformId) => state.ids[platformId],
    setGraphIds(entries) {
      Object.assign(state.ids, entries);
    },
    overdue: (platformId) => state.overdue[platformId],
    setOverdue(platformId, taskIds) {
      state.overdue[platformId] = taskIds;
    },
    async save() {
      this.saves += 1;
    },
    state,
  };
}

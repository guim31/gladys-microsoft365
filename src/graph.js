// -----------------------------------------------------------------------------
// Microsoft Graph client: the few calls the integration makes, one at a time.
//
// Throttling (https://learn.microsoft.com/graph/throttling): Outlook allows
// 10,000 requests per 10 minutes and 4 concurrent requests per application
// and mailbox; beyond, Graph answers 429 (or 503/504) with `Retry-After`.
// The client honours it: a short wait is waited (twice at most), a long one
// blocks every call until it is over, without asking Graph again. Requests
// are serialized, so the integration never holds more than one at a time.
// -----------------------------------------------------------------------------

import { AuthError } from './auth.js';
import { graphDateTime, graphMidnight } from './time.js';

export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const RETRYABLE = new Set([429, 503, 504]);
// Waits up to this long are waited inside the request; longer ones block.
const MAX_INLINE_WAIT_MS = 30_000;
const MAX_PAGES = 20;

export class GraphError extends Error {
  /**
   * @param {number} status HTTP status (0: no answer)
   * @param {string} code Graph error code (`ErrorItemNotFound`…) or ours
   * @param {{ retryAt?: number }} [details]
   */
  constructor(status, code, { retryAt } = {}) {
    super(status ? `HTTP ${status} ${code}` : code);
    this.name = 'GraphError';
    this.status = status;
    this.code = code;
    this.retryAt = retryAt;
  }

  get throttled() {
    return this.retryAt !== undefined;
  }
}

/** Retry-After: seconds, or an HTTP date. */
function retryAfterMs(header, now) {
  if (!header) {
    return undefined;
  }
  const seconds = Number(header);
  if (Number.isFinite(seconds)) {
    return Math.max(0, seconds * 1000);
  }
  const date = Date.parse(header);
  return Number.isNaN(date) ? undefined : Math.max(0, date - now);
}

const defaultSleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * @param {object} options
 * @param {ReturnType<import('./auth.js').createAuth>} options.auth
 * @param {() => object} options.getConfig the current normalized config
 * @param {typeof fetch} [options.fetch]
 * @param {() => number} [options.now]
 * @param {(ms: number) => Promise<void>} [options.sleep]
 * @param {number} [options.timeoutMs]
 */
export function createGraphClient({
  auth,
  getConfig,
  fetch = globalThis.fetch,
  now = Date.now,
  sleep = defaultSleep,
  timeoutMs = 20_000,
}) {
  let blockedUntil = 0;
  let queue = Promise.resolve();

  async function send(method, url, { body, headers = {} } = {}, attempt = 0, renewed = false) {
    if (now() < blockedUntil) {
      throw new GraphError(429, 'throttled', { retryAt: blockedUntil });
    }
    const config = getConfig();
    const token = await auth.accessToken(config, { force: renewed });
    let response;
    try {
      response = await fetch(url, {
        method,
        headers: {
          authorization: `Bearer ${token}`,
          accept: 'application/json',
          ...(body === undefined ? {} : { 'content-type': 'application/json' }),
          ...headers,
        },
        body: body === undefined ? undefined : JSON.stringify(body),
        signal: AbortSignal.timeout(timeoutMs),
      });
    } catch (err) {
      throw new GraphError(0, err?.name === 'TimeoutError' ? 'timeout' : 'network_error');
    }

    if (response.ok) {
      if (response.status === 204) {
        return null;
      }
      return response.json().catch(() => null);
    }

    let code = `http_${response.status}`;
    try {
      const payload = await response.json();
      if (typeof payload?.error?.code === 'string') {
        code = payload.error.code;
      }
    } catch {
      // No JSON body: keep the status.
    }

    // An access token refused before its time (revoked, clock skew): renew once.
    if (response.status === 401 && !renewed) {
      return send(method, url, { body, headers }, attempt, true);
    }
    if (RETRYABLE.has(response.status)) {
      const wait = retryAfterMs(response.headers.get('retry-after'), now()) ?? 2 ** attempt * 2000;
      if (wait <= MAX_INLINE_WAIT_MS && attempt < 2) {
        await sleep(wait);
        return send(method, url, { body, headers }, attempt + 1, renewed);
      }
      blockedUntil = now() + Math.max(wait, 1000);
      throw new GraphError(response.status, code, { retryAt: blockedUntil });
    }
    throw new GraphError(response.status, code);
  }

  /** Serialized: one request at a time, in call order. */
  function request(method, path, options = {}) {
    const url = path.startsWith('https://') ? path : `${GRAPH_BASE}${path}`;
    const run = queue.then(() => send(method, url, options));
    queue = run.catch(() => {});
    return run;
  }

  /** Every page of a collection (`@odata.nextLink`), Graph links only. */
  async function getAll(path, options = {}) {
    const items = [];
    let next = path;
    for (let page = 0; next && page < MAX_PAGES; page += 1) {
      const answer = await request('GET', next, options);
      items.push(...(Array.isArray(answer?.value) ? answer.value : []));
      const link = answer?.['@odata.nextLink'];
      next = typeof link === 'string' && link.startsWith(`${GRAPH_BASE}/`) ? link : null;
    }
    return items;
  }

  const id = (value) => encodeURIComponent(value);
  const UTC = { prefer: 'outlook.timezone="UTC"' };

  return {
    request,
    getAll,

    /** The signed-in user: `{ id, displayName }`. */
    me() {
      return request('GET', '/me?$select=id,displayName');
    },

    calendars() {
      return getAll('/me/calendars?$select=id,name,isDefaultCalendar,canEdit&$top=100');
    },

    /** The occurrences overlapping [from, to), recurrences expanded, in UTC. */
    calendarView(calendarId, fromMs, toMs) {
      const query = new URLSearchParams({
        startDateTime: new Date(fromMs).toISOString(),
        endDateTime: new Date(toMs).toISOString(),
        $select: 'id,subject,start,end,isAllDay,showAs,isCancelled,location,responseStatus',
        $orderby: 'start/dateTime',
        $top: '200',
      });
      return getAll(`/me/calendars/${id(calendarId)}/calendarView?${query}`, { headers: UTC });
    },

    async createEvent(calendarId, { subject, startMs, endMs }) {
      return request('POST', `/me/calendars/${id(calendarId)}/events`, {
        body: {
          subject,
          start: graphDateTime(startMs),
          end: graphDateTime(endMs),
          showAs: 'busy',
        },
        headers: UTC,
      });
    },

    todoLists() {
      return getAll('/me/todo/lists?$top=100');
    },

    /** The tasks of a list that are not completed. */
    openTasks(listId) {
      const query = new URLSearchParams({
        $filter: "status ne 'completed'",
        $select: 'id,title,status,dueDateTime,importance,createdDateTime',
        $top: '100',
      });
      return getAll(`/me/todo/lists/${id(listId)}/tasks?${query}`);
    },

    /** @param {{ title: string, dueDate?: string }} task dueDate 'YYYY-MM-DD' */
    createTask(listId, { title, dueDate }) {
      return request('POST', `/me/todo/lists/${id(listId)}/tasks`, {
        body: { title, ...(dueDate ? { dueDateTime: graphMidnight(dueDate) } : {}) },
      });
    },

    completeTask(listId, taskId) {
      return request('PATCH', `/me/todo/lists/${id(listId)}/tasks/${id(taskId)}`, {
        body: { status: 'completed' },
      });
    },

    /** When Graph throttles every call (ms timestamp), or 0. */
    get blockedUntil() {
      return now() < blockedUntil ? blockedUntil : 0;
    },
  };
}

/** True when the error means the session is gone (sign in again). */
export function needsSignIn(err) {
  return err instanceof AuthError && err.needsSignIn;
}

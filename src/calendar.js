// -----------------------------------------------------------------------------
// Outlook calendars: from the occurrences Graph returns to what Gladys shows
// (busy, out of office, current and next event) and to the scene triggers
// (an event starts, ends, starts in N minutes).
//
// The triggers are computed from the occurrences already read, every 30
// seconds, not from Graph: an event starting at 10:00 fires at 10:00 even if
// Graph is read every 15 minutes. Each check looks at the instants between
// the previous check and now, so an instant is never announced twice — and
// never after a restart (the first check only takes note of the time).
// -----------------------------------------------------------------------------

import { fold } from './config.js';
import { MINUTE_MS, graphDate, graphInstant, localMidnight } from './time.js';

/** "In N minutes" choices of the `event_upcoming` trigger. */
export const UPCOMING_MINUTES = [5, 10, 15, 30, 60, 120];

// `showAs` values that make the calendar busy (free and workingElsewhere don't).
const BUSY = new Set(['busy', 'oof', 'tentative']);

// Past this delay, an instant missed by the checks (container paused, Gladys
// unreachable) is no longer announced: a scene must not run an hour late.
export const MAX_LATENESS_MS = 5 * MINUTE_MS;

/**
 * One Graph occurrence, as the integration uses it; null when cancelled,
 * declined or unreadable.
 * @returns {{ id, title, location, startMs, endMs, allDay, showAs } | null}
 */
export function normalizeEvent(raw) {
  if (!raw || typeof raw !== 'object' || raw.isCancelled === true) {
    return null;
  }
  if (raw.responseStatus?.response === 'declined') {
    return null;
  }
  const allDay = raw.isAllDay === true;
  let startMs;
  let endMs;
  if (allDay) {
    // A date, not an instant: from the local midnight of its first day to
    // the one of the day after its last (Graph's end is exclusive).
    const startDate = graphDate(raw.start);
    const endDate = graphDate(raw.end);
    if (!startDate) {
      return null;
    }
    startMs = localMidnight(startDate);
    endMs = endDate && endDate > startDate ? localMidnight(endDate) : localMidnight(startDate, 1);
  } else {
    startMs = graphInstant(raw.start);
    endMs = graphInstant(raw.end);
    if (Number.isNaN(startMs)) {
      return null;
    }
    if (Number.isNaN(endMs) || endMs < startMs) {
      endMs = startMs;
    }
  }
  return {
    id: String(raw.id ?? ''),
    title: String(raw.subject ?? '').trim(),
    location: String(raw.location?.displayName ?? '').trim(),
    startMs,
    endMs,
    allDay,
    showAs: String(raw.showAs ?? 'busy'),
  };
}

/** The normalized occurrences, sorted by start. */
export function normalizeEvents(rawEvents) {
  return rawEvents
    .map(normalizeEvent)
    .filter(Boolean)
    .sort((a, b) => a.startMs - b.startMs || a.endMs - b.endMs);
}

// A zero-duration event (a reminder) is never "in progress".
const inProgress = (event, now) => event.startMs <= now && now < event.endMs;

/**
 * What the calendar device shows at `now`.
 * @returns {{ busy: boolean, outOfOffice: boolean, current: object|null, next: object|null }}
 */
export function calendarState(events, now) {
  const running = events.filter((event) => inProgress(event, now));
  // The most telling current event: a timed one before an all-day one, a
  // busy one before a free one, then the one that started last.
  const current =
    [...running].sort(
      (a, b) =>
        Number(a.allDay) - Number(b.allDay) ||
        Number(BUSY.has(b.showAs)) - Number(BUSY.has(a.showAs)) ||
        b.startMs - a.startMs,
    )[0] ?? null;
  return {
    busy: running.some((event) => BUSY.has(event.showAs)),
    outOfOffice: running.some((event) => event.showAs === 'oof'),
    current,
    next: events.find((event) => event.startMs > now) ?? null,
  };
}

/** The first watched keyword the title contains (case and accents ignored). */
export function matchKeyword(title, keywords) {
  const folded = fold(title);
  return keywords.find((keyword) => folded.includes(fold(keyword))) ?? null;
}

/**
 * The trigger instants in (from, to]: starts, ends and "in N minutes".
 * @returns {Array<{ key: 'event_started'|'event_ended'|'event_upcoming', event, at: number, minutes?: number }>}
 */
export function transitions(events, from, to) {
  const found = [];
  const within = (at) => at > from && at <= to;
  for (const event of events) {
    if (within(event.startMs)) {
      found.push({ key: 'event_started', event, at: event.startMs });
    }
    if (event.endMs > event.startMs && within(event.endMs)) {
      found.push({ key: 'event_ended', event, at: event.endMs });
    }
    // All-day events have no "in N minutes": the evening before is not a lead time.
    if (!event.allDay) {
      for (const minutes of UPCOMING_MINUTES) {
        const at = event.startMs - minutes * MINUTE_MS;
        if (within(at)) {
          found.push({ key: 'event_upcoming', event, at, minutes });
        }
      }
    }
  }
  return found.sort((a, b) => a.at - b.at);
}

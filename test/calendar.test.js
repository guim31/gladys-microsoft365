// -----------------------------------------------------------------------------
// Outlook occurrences: reading Graph's dates, the states of a calendar device
// and the trigger instants.
// -----------------------------------------------------------------------------

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { NOW, fixture } from './helpers/world.js';
import {
  calendarState,
  matchKeyword,
  normalizeEvent,
  normalizeEvents,
  transitions,
} from '../src/calendar.js';
import { graphDate, graphInstant, graphMidnight } from '../src/time.js';

const MINUTE = 60_000;
const events = normalizeEvents(fixture('graph/calendar-view.json').value);
const titled = (title) => events.find((event) => event.title === title);

test('Graph instants: UTC, an IANA zone, a Windows name read as UTC', () => {
  assert.equal(
    graphInstant({ dateTime: '2026-10-12T07:30:00.0000000', timeZone: 'UTC' }),
    Date.parse('2026-10-12T07:30:00Z'),
  );
  assert.equal(
    graphInstant({ dateTime: '2026-10-12T09:30:00', timeZone: 'Europe/Paris' }),
    Date.parse('2026-10-12T07:30:00Z'),
  );
  // Winter time in Paris, and New York in summer.
  assert.equal(
    graphInstant({ dateTime: '2026-12-01T09:30:00', timeZone: 'Europe/Paris' }),
    Date.parse('2026-12-01T08:30:00Z'),
  );
  assert.equal(
    graphInstant({ dateTime: '2026-07-01T09:30:00', timeZone: 'America/New_York' }),
    Date.parse('2026-07-01T13:30:00Z'),
  );
  assert.equal(
    graphInstant({ dateTime: '2026-10-12T07:30:00', timeZone: 'W. Europe Standard Time' }),
    Date.parse('2026-10-12T07:30:00Z'),
  );
  assert.ok(Number.isNaN(graphInstant({ dateTime: 'tomorrow', timeZone: 'UTC' })));
});

test('Graph dates: the nearest midnight, whatever zone it was written in', () => {
  // The To Do example of the Graph reference: midnight in New York, in UTC.
  assert.equal(
    graphDate({ dateTime: '2020-08-25T04:00:00.0000000', timeZone: 'UTC' }),
    '2020-08-25',
  );
  // Midnight in Paris, in UTC (the day before, 22:00).
  assert.equal(
    graphDate({ dateTime: '2026-10-11T22:00:00.0000000', timeZone: 'UTC' }),
    '2026-10-12',
  );
  // Midnight as written, and midnight in Tokyo.
  assert.equal(graphDate({ dateTime: '2026-10-12T00:00:00', timeZone: 'UTC' }), '2026-10-12');
  assert.equal(graphDate({ dateTime: '2026-10-11T15:00:00', timeZone: 'UTC' }), '2026-10-12');
  // What the integration writes reads back as the same date.
  assert.equal(graphDate(graphMidnight('2026-10-20')), '2026-10-20');
  assert.deepEqual(graphMidnight('2026-10-20'), {
    dateTime: '2026-10-19T22:00:00',
    timeZone: 'UTC',
  });
});

test('cancelled and declined occurrences are left out', () => {
  assert.deepEqual(
    events.map((event) => event.title),
    ['Télétravail', "Point d'équipe", 'Dentiste', 'Déjeuner avec Léa'],
  );
  assert.equal(normalizeEvent({ isCancelled: true }), null);
  assert.equal(normalizeEvent(null), null);
});

test('an all-day event covers its local day, from midnight to midnight', () => {
  const remote = titled('Télétravail');
  assert.equal(remote.allDay, true);
  assert.equal(remote.startMs, Date.parse('2026-10-11T22:00:00Z'));
  assert.equal(remote.endMs, Date.parse('2026-10-12T22:00:00Z'));
});

test('busy, out of office, current and next event', () => {
  // 08:00: only the remote-work day (working elsewhere: not busy).
  let state = calendarState(events, NOW);
  assert.equal(state.busy, false);
  assert.equal(state.outOfOffice, false);
  assert.equal(state.current.title, 'Télétravail');
  assert.equal(state.next.title, "Point d'équipe");

  // 09:45: the meeting, timed, is the current event.
  state = calendarState(events, Date.parse('2026-10-12T07:45:00Z'));
  assert.equal(state.busy, true);
  assert.equal(state.current.title, "Point d'équipe");
  assert.equal(state.next.title, 'Dentiste');

  // 14:15: out of office at the dentist.
  state = calendarState(events, Date.parse('2026-10-12T12:15:00Z'));
  assert.equal(state.busy, true);
  assert.equal(state.outOfOffice, true);

  // 10:00 sharp: the meeting is over (the end is exclusive).
  state = calendarState(events, Date.parse('2026-10-12T08:00:00Z'));
  assert.equal(state.busy, false);

  // Tomorrow: a tentative lunch counts as busy.
  state = calendarState(events, Date.parse('2026-10-13T11:00:00Z'));
  assert.equal(state.busy, true);
  assert.equal(state.current.title, 'Déjeuner avec Léa');
  assert.equal(state.next, null);
});

test('keywords: case and accents ignored, the first listed wins', () => {
  assert.equal(matchKeyword('Télétravail', ['dentiste', 'teletravail']), 'teletravail');
  assert.equal(matchKeyword('RDV DENTISTE', ['Dentiste', 'RDV']), 'Dentiste');
  assert.equal(matchKeyword('Réunion', ['Dentiste']), null);
  assert.equal(matchKeyword('Réunion', []), null);
});

test('trigger instants: starts, ends, and every lead time once', () => {
  const meeting = titled("Point d'équipe");
  const found = transitions(events, NOW, Date.parse('2026-10-12T09:00:00Z'));
  const of = (event) =>
    found.filter((entry) => entry.event === event).map((entry) => [entry.key, entry.minutes]);
  assert.deepEqual(of(meeting), [
    ['event_upcoming', 60],
    ['event_upcoming', 30],
    ['event_upcoming', 15],
    ['event_upcoming', 10],
    ['event_upcoming', 5],
    ['event_started', undefined],
    ['event_ended', undefined],
  ]);
  // The window is (from, to]: an instant equal to `from` was the previous one's.
  assert.deepEqual(
    transitions(events, meeting.startMs, meeting.startMs + MINUTE).map((entry) => entry.key),
    [],
  );
  assert.deepEqual(
    transitions(events, meeting.startMs - MINUTE, meeting.startMs).map((entry) => entry.key),
    ['event_started'],
  );
});

test('all-day events start and end, but have no lead time', () => {
  const remote = titled('Télétravail');
  const found = transitions([remote], remote.startMs - 3 * 3_600_000, remote.endMs);
  assert.deepEqual(
    found.map((entry) => entry.key),
    ['event_started', 'event_ended'],
  );
});

// -----------------------------------------------------------------------------
// Dates and times.
//
// Calendar dates (a due date, an all-day event) are read in the time zone of
// the process, which the Gladys supervisor sets to the one of the instance
// (TZ). Instants from Microsoft Graph are requested in UTC
// (`Prefer: outlook.timezone="UTC"`), but a `dateTimeTimeZone` can still name
// another zone: an IANA name is converted, anything else is read as UTC.
// -----------------------------------------------------------------------------

export const MINUTE_MS = 60_000;
export const HOUR_MS = 3_600_000;
export const DAY_MS = 86_400_000;

const UTC_NAMES = new Set(['', 'utc', 'etc/utc', 'gmt', 'etc/gmt', 'coordinated universal time']);

const pad = (number) => String(number).padStart(2, '0');

/** 'YYYY-MM-DD' of an instant, in the local time zone. */
export function localDateKey(ms) {
  const date = new Date(ms);
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

/** Local midnight of a 'YYYY-MM-DD' date, plus `days` calendar days. */
export function localMidnight(dateKey, days = 0) {
  const [year, month, day] = dateKey.split('-').map(Number);
  return new Date(year, month - 1, day + days).getTime();
}

/** The calendar day `days` after (or before) a 'YYYY-MM-DD' date. */
export function addDays(dateKey, days) {
  const [year, month, day] = dateKey.split('-').map(Number);
  const date = new Date(Date.UTC(year, month - 1, day + days));
  return date.toISOString().slice(0, 10);
}

function isIanaZone(zone) {
  try {
    new Intl.DateTimeFormat('en-US', { timeZone: zone });
    return zone.includes('/');
  } catch {
    return false;
  }
}

/** Offset of `zone` from UTC at an instant, in ms (positive east of Greenwich). */
function zoneOffsetMs(zone, ms) {
  const parts = Object.fromEntries(
    new Intl.DateTimeFormat('en-US', {
      timeZone: zone,
      hourCycle: 'h23',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
    })
      .formatToParts(ms)
      .map((part) => [part.type, part.value]),
  );
  const asUtc = Date.UTC(
    Number(parts.year),
    Number(parts.month) - 1,
    Number(parts.day),
    Number(parts.hour),
    Number(parts.minute),
    Number(parts.second),
  );
  return asUtc - Math.floor(ms / 1000) * 1000;
}

/**
 * The instant of a Graph `dateTimeTimeZone` (`{ dateTime, timeZone }`), in ms;
 * NaN when unreadable.
 */
export function graphInstant(value) {
  const dateTime = String(value?.dateTime ?? '');
  const match = dateTime.match(/^(\d{4}-\d{2}-\d{2})(?:T(\d{2}:\d{2})(?::(\d{2}))?(?:\.\d+)?)?$/);
  if (!match) {
    return NaN;
  }
  const iso = `${match[1]}T${match[2] ?? '00:00'}:${match[3] ?? '00'}`;
  const asUtc = Date.parse(`${iso}Z`);
  const zone = String(value?.timeZone ?? '').trim();
  if (UTC_NAMES.has(zone.toLowerCase()) || !isIanaZone(zone)) {
    return asUtc;
  }
  // Two passes: the offset at the guessed instant, then at the corrected one
  // (right on both sides of a daylight saving change).
  const first = asUtc - zoneOffsetMs(zone, asUtc);
  return asUtc - zoneOffsetMs(zone, first);
}

/**
 * The calendar date ('YYYY-MM-DD') a Graph `dateTimeTimeZone` stands for, for
 * the values that are dates: a To Do due date, an all-day event. Microsoft
 * stores them as a midnight of some time zone, often converted to UTC
 * (the midnight of Paris reads "the day before, 22:00 UTC"): the nearest
 * midnight is the date, whatever the zone between UTC-11 and UTC+12.
 */
export function graphDate(value) {
  const ms = graphInstant(value);
  if (Number.isNaN(ms)) {
    return null;
  }
  return new Date(ms + 12 * HOUR_MS).toISOString().slice(0, 10);
}

/** The Graph `dateTimeTimeZone` of the local midnight of a date, in UTC. */
export function graphMidnight(dateKey) {
  return {
    dateTime: new Date(localMidnight(dateKey)).toISOString().slice(0, 19),
    timeZone: 'UTC',
  };
}

/** The Graph `dateTimeTimeZone` of an instant, in UTC. */
export function graphDateTime(ms) {
  return { dateTime: new Date(ms).toISOString().slice(0, 19), timeZone: 'UTC' };
}

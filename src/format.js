// -----------------------------------------------------------------------------
// Dates and events as people read them, in the language of the setting (or
// of the dashboard, for the widget). Local time zone of the process (TZ).
// -----------------------------------------------------------------------------

import { LOCALES, TEXT, pick } from './i18n.js';
import { addDays, localDateKey } from './time.js';

const locale = (language) => LOCALES[language] ?? LOCALES.en;

/** "14:00" (or "2:00 pm" where the language says so). */
export function formatTime(ms, language) {
  return new Intl.DateTimeFormat(locale(language), { hour: '2-digit', minute: '2-digit' }).format(
    ms,
  );
}

/** "Thu 16 Oct", in the language. */
export function formatDay(ms, language) {
  return new Intl.DateTimeFormat(locale(language), {
    weekday: 'short',
    day: 'numeric',
    month: 'short',
  }).format(ms);
}

/** "today", "tomorrow" or "Thu 16 Oct" for a day, relative to `now`. */
export function formatRelativeDay(ms, language, now) {
  const text = pick(TEXT, language);
  const day = localDateKey(ms);
  const today = localDateKey(now);
  if (day === today) {
    return text.today;
  }
  if (day === addDays(today, 1)) {
    return text.tomorrow;
  }
  return formatDay(ms, language);
}

/** When an event starts: "14:00" today, "tomorrow 14:00", "Thu 16 Oct" for an all-day one. */
export function formatStart(event, language, now) {
  const day = formatRelativeDay(event.startMs, language, now);
  if (event.allDay) {
    return day;
  }
  if (localDateKey(event.startMs) === localDateKey(now)) {
    return formatTime(event.startMs, language);
  }
  return `${day} ${formatTime(event.startMs, language)}`;
}

/** The time range of an event of today: "09:00 – 10:30", or "all day". */
export function formatRange(event, language) {
  if (event.allDay) {
    return pick(TEXT, language).allDay;
  }
  if (event.endMs <= event.startMs) {
    return formatTime(event.startMs, language);
  }
  return `${formatTime(event.startMs, language)} – ${formatTime(event.endMs, language)}`;
}

/** A due date ('YYYY-MM-DD') relative to today: "today", "tomorrow", "Thu 16 Oct". */
export function formatDueDate(dueDate, language, now) {
  const [year, month, day] = dueDate.split('-').map(Number);
  return formatRelativeDay(new Date(year, month - 1, day, 12).getTime(), language, now);
}

/** Cut a text to `max` characters, with an ellipsis. */
export function truncate(text, max) {
  const value = String(text ?? '');
  return value.length <= max ? value : `${value.slice(0, max - 1).trimEnd()}…`;
}

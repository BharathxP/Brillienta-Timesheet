// Calendar-day helpers. Days are handled as yyyy-mm-dd keys in local time, never as instants, so a week never shifts
// when people work in different time zones. Pure functions (tests/dates.test.cjs).

import { DateKey } from '../models/types';

export const DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
export const DAY_SHORT = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
export const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

const pad = (n: number): string => (n < 10 ? '0' : '') + n;

export function toKey(d: Date): DateKey {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/** Local midnight of a key (or of the date part of an ISO string). */
export function fromKey(key: DateKey): Date {
  const [y, m, d] = (key || '').substring(0, 10).split('-').map(Number);
  return new Date(y || 1970, (m || 1) - 1, d || 1);
}

export function isKey(s: string): boolean {
  return /^\d{4}-\d{2}-\d{2}$/.test(s || '') && !isNaN(fromKey(s).getTime());
}

export function addDays(key: DateKey, days: number): DateKey {
  const d = fromKey(key);
  d.setDate(d.getDate() + days);
  return toKey(d);
}

export function dayOfWeek(key: DateKey): number {
  return fromKey(key).getDay();
}

/** First day of the week containing `key`, for weeks starting on `startDay` (0 = Sunday). */
export function weekStartOf(key: DateKey, startDay: number): DateKey {
  const diff = (dayOfWeek(key) - startDay + 7) % 7;
  return addDays(key, -diff);
}

/** The seven days of the week that starts on `weekStart`. */
export function weekDays(weekStart: DateKey): DateKey[] {
  const out: DateKey[] = [];
  for (let i = 0; i < 7; i++) out.push(addDays(weekStart, i));
  return out;
}

export function daysBetween(a: DateKey, b: DateKey): number {
  return Math.round((fromKey(b).getTime() - fromKey(a).getTime()) / 86400000);
}

export function today(): DateKey {
  return toKey(new Date());
}

export function monthKey(key: DateKey): string {
  return key.substring(0, 7);
}

export function formatDay(key: DateKey, withYear: boolean = false): string {
  if (!key) return '';
  const d = fromKey(key);
  return `${d.getDate()} ${MONTHS[d.getMonth()]}${withYear ? ' ' + d.getFullYear() : ''}`;
}

export function formatWeek(weekStart: DateKey): string {
  const end = addDays(weekStart, 6);
  const a = fromKey(weekStart);
  const b = fromKey(end);
  if (a.getFullYear() !== b.getFullYear()) return `${formatDay(weekStart, true)} – ${formatDay(end, true)}`;
  if (a.getMonth() === b.getMonth()) return `${a.getDate()} – ${b.getDate()} ${MONTHS[b.getMonth()]} ${b.getFullYear()}`;
  return `${formatDay(weekStart)} – ${formatDay(end)} ${b.getFullYear()}`;
}

export function formatMonth(month: string): string {
  const [y, m] = month.split('-').map(Number);
  return `${MONTHS[(m || 1) - 1]} ${y}`;
}

/** ISO-8601 week number (for labels). */
export function isoWeek(key: DateKey): number {
  const d = fromKey(key);
  d.setDate(d.getDate() + 3 - ((d.getDay() + 6) % 7));
  const jan4 = new Date(d.getFullYear(), 0, 4);
  return 1 + Math.round(((d.getTime() - jan4.getTime()) / 86400000 - 3 + ((jan4.getDay() + 6) % 7)) / 7);
}

/** A date range for a named period, relative to `ref`. */
export function periodRange(period: string, ref: DateKey, weekStart: number): { from: DateKey; to: DateKey } {
  const d = fromKey(ref);
  const y = d.getFullYear();
  const m = d.getMonth();
  const monthStart = (yy: number, mm: number): DateKey => toKey(new Date(yy, mm, 1));
  const monthEnd = (yy: number, mm: number): DateKey => toKey(new Date(yy, mm + 1, 0));
  switch (period) {
    case 'thisWeek': {
      const s = weekStartOf(ref, weekStart);
      return { from: s, to: addDays(s, 6) };
    }
    case 'lastWeek': {
      const s = addDays(weekStartOf(ref, weekStart), -7);
      return { from: s, to: addDays(s, 6) };
    }
    case 'thisMonth':
      return { from: monthStart(y, m), to: monthEnd(y, m) };
    case 'lastMonth':
      return { from: monthStart(y, m - 1), to: monthEnd(y, m - 1) };
    case 'thisQuarter': {
      const q = Math.floor(m / 3) * 3;
      return { from: monthStart(y, q), to: monthEnd(y, q + 2) };
    }
    case 'thisYear':
      return { from: `${y}-01-01`, to: `${y}-12-31` };
    case 'last30':
      return { from: addDays(ref, -29), to: ref };
    case 'last90':
      return { from: addDays(ref, -89), to: ref };
    default:
      return { from: monthStart(y, m), to: monthEnd(y, m) };
  }
}

/** Week starts (keys) whose weeks overlap [from, to]. */
export function weeksIn(from: DateKey, to: DateKey, weekStart: number): DateKey[] {
  const out: DateKey[] = [];
  let w = weekStartOf(from, weekStart);
  let guard = 0;
  while (w <= to && guard++ < 600) {
    out.push(w);
    w = addDays(w, 7);
  }
  return out;
}

export const PERIODS: Array<{ key: string; text: string }> = [
  { key: 'thisWeek', text: 'This week' },
  { key: 'lastWeek', text: 'Last week' },
  { key: 'thisMonth', text: 'This month' },
  { key: 'lastMonth', text: 'Last month' },
  { key: 'thisQuarter', text: 'This quarter' },
  { key: 'thisYear', text: 'This year' },
  { key: 'last30', text: 'Last 30 days' },
  { key: 'last90', text: 'Last 90 days' }
];

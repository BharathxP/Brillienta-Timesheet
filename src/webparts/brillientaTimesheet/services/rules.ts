// Which days a person tracks time on, which hour limits apply to them, and whether a week passes those limits.
// Pure functions (tests/rules.test.cjs).

import { DateKey, IHourLimits, IHourRule, ITimeEntry, ITimeOff, ITimesheetConfig, ITsUser } from '../models/types';
import { dayOfWeek, formatDay } from './dates';

export interface IEffectiveLimits extends IHourLimits {
  enforcement: 'block' | 'warn';
  allowNonWorkingDays: boolean;
  /** Where the limits come from, for the screen ("Organisation limits", "Rule: Contractors"). */
  source: string;
}

export interface IIssue {
  level: 'error' | 'warning';
  message: string;
  day?: DateKey;
}

/** Days of the week this person tracks time on. */
export function workDaysFor(cfg: ITimesheetConfig, user?: ITsUser): number[] {
  return user && user.workDays && user.workDays.length ? user.workDays : cfg.workDays;
}

export function holidayOn(cfg: ITimesheetConfig, day: DateKey): string | undefined {
  const h = cfg.holidays.filter((x) => x.date === day)[0];
  return h ? h.name || 'Holiday' : undefined;
}

/** A day time is normally tracked on: a working day that is not a holiday. */
export function isTrackingDay(cfg: ITimesheetConfig, day: DateKey, user?: ITsUser): boolean {
  return workDaysFor(cfg, user).indexOf(dayOfWeek(day)) >= 0 && !holidayOn(cfg, day);
}

/** The first hour rule (in order) whose conditions all match the person. */
export function matchRule(rules: IHourRule[], user: ITsUser): IHourRule | undefined {
  if (user.ruleId) {
    const pinned = rules.filter((r) => r.id === user.ruleId)[0];
    if (pinned) return pinned;
  }
  const lc = (s: string): string => (s || '').toLowerCase();
  return rules.filter(
    (r) =>
      (!r.roles.length || r.roles.indexOf(user.role) >= 0) &&
      (!r.departments.length || r.departments.some((d) => lc(d) === lc(user.department))) &&
      (!r.emails.length || r.emails.some((e) => lc(e) === lc(user.email))) &&
      (!r.externalOnly || user.external)
  )[0];
}

/** The limits that apply to a person: their rule (rule-based restrictions on), else the organisation's. */
export function limitsFor(cfg: ITimesheetConfig, user?: ITsUser): IEffectiveLimits {
  const rule = cfg.rulesEnabled && user ? matchRule(cfg.rules, user) : undefined;
  if (rule) {
    return {
      dailyMin: rule.dailyMin,
      dailyMax: rule.dailyMax,
      weeklyMin: rule.weeklyMin,
      weeklyMax: rule.weeklyMax,
      enforcement: rule.enforcement,
      allowNonWorkingDays: rule.allowNonWorkingDays,
      source: `Rule: ${rule.name}`
    };
  }
  const l = cfg.limits;
  return { dailyMin: l.dailyMin, dailyMax: l.dailyMax, weeklyMin: l.weeklyMin, weeklyMax: l.weeklyMax, enforcement: l.enforcement, allowNonWorkingDays: true, source: 'Organisation limits' };
}

const round2 = (n: number): number => Math.round(n * 100) / 100;

export function dayTotals(entries: ITimeEntry[], days: DateKey[]): Record<DateKey, number> {
  const out: Record<DateKey, number> = {};
  days.forEach((d) => {
    out[d] = round2(entries.reduce((s, e) => s + (Number(e.hours[d]) || 0), 0));
  });
  return out;
}

/** Hours of approved leave on a day for a person (a full standard day unless the leave says otherwise). */
export function leaveHours(timeOff: ITimeOff[], day: DateKey, standardDay: number): number {
  const l = timeOff.filter((t) => t.start <= day && t.end >= day)[0];
  return l ? (l.hoursPerDay !== undefined ? l.hoursPerDay : standardDay) : 0;
}

export interface IValidateInput {
  cfg: ITimesheetConfig;
  user?: ITsUser;
  days: DateKey[];
  entries: ITimeEntry[];
  /** Approved time off (counts towards the minimums when it is not already on the sheet). */
  timeOff?: ITimeOff[];
  /** Entry ids on the time-off activity (already counted). */
  timeOffActivityId?: number;
  /** Check the "required" settings that only matter on submit. */
  forSubmit?: boolean;
  /** Activity tracking is on and an activity is required on each line. */
  requireActivity?: boolean;
  /** Required custom columns (entry level). */
  requiredColumns?: Array<{ key: string; label: string }>;
}

/** Problems with a week. Errors block submitting when the limits are enforced ('block'); warnings never block. */
export function validateWeek(input: IValidateInput): IIssue[] {
  const { cfg, user, days, entries } = input;
  const limits = limitsFor(cfg, user);
  const issues: IIssue[] = [];
  const level: IIssue['level'] = limits.enforcement === 'block' ? 'error' : 'warning';
  const totals = dayTotals(entries, days);
  const hasLeaveOnSheet = (d: DateKey): boolean => !!input.timeOffActivityId && entries.some((e) => e.activityId === input.timeOffActivityId && (e.hours[d] || 0) > 0);
  let week = 0;
  days.forEach((d) => {
    const leave = input.timeOff && !hasLeaveOnSheet(d) ? leaveHours(input.timeOff, d, cfg.standardDay) : 0;
    const t = totals[d] + leave;
    week += t;
    const tracking = isTrackingDay(cfg, d, user);
    if (totals[d] < 0) issues.push({ level: 'error', day: d, message: `${formatDay(d)}: hours cannot be negative.` });
    if (totals[d] > 24) issues.push({ level: 'error', day: d, message: `${formatDay(d)}: more than 24 hours in one day.` });
    if (!tracking && totals[d] > 0 && !limits.allowNonWorkingDays) {
      const why = holidayOn(cfg, d) ? `is a holiday (${holidayOn(cfg, d)})` : 'is not a working day';
      issues.push({ level, day: d, message: `${formatDay(d)} ${why}; time cannot be logged on it.` });
    }
    if (limits.dailyMax > 0 && t > limits.dailyMax) issues.push({ level, day: d, message: `${formatDay(d)}: ${round2(t)} h is more than the daily maximum of ${limits.dailyMax} h.` });
    if (input.forSubmit && tracking && limits.dailyMin > 0 && t < limits.dailyMin) {
      issues.push({ level, day: d, message: `${formatDay(d)}: ${round2(t)} h is less than the daily minimum of ${limits.dailyMin} h.` });
    }
    if (input.forSubmit && cfg.requireNotes) {
      entries.forEach((e) => {
        if ((e.hours[d] || 0) > 0 && !(e.notes[d] || '').trim()) issues.push({ level: 'error', day: d, message: `${formatDay(d)}: add a note for every entry with hours.` });
      });
    }
  });
  week = round2(week);
  if (limits.weeklyMax > 0 && week > limits.weeklyMax) issues.push({ level, message: `${week} h this week is more than the weekly maximum of ${limits.weeklyMax} h.` });
  if (input.forSubmit && limits.weeklyMin > 0 && week < limits.weeklyMin) issues.push({ level, message: `${week} h this week is less than the weekly minimum of ${limits.weeklyMin} h.` });
  if (input.forSubmit) {
    if (!entries.some((e) => e.total > 0) && week === 0) issues.push({ level: 'error', message: 'The week has no hours.' });
    entries.forEach((e) => {
      if (e.total > 0 && !e.projectId && !(input.timeOffActivityId && e.activityId === input.timeOffActivityId)) issues.push({ level: 'error', message: 'Every line with hours needs a project.' });
      if (e.total > 0 && input.requireActivity && !e.activityId) issues.push({ level: 'error', message: 'Every line with hours needs an activity.' });
      (input.requiredColumns || []).forEach((c) => {
        const v = e.custom[c.key];
        if (e.total > 0 && (v === undefined || v === '' || v === null)) issues.push({ level: 'error', message: `Fill in "${c.label}" on every line with hours.` });
      });
    });
  }
  // One message per text (a missing note on five lines is one problem to fix).
  const seen: Record<string, boolean> = {};
  return issues.filter((i) => (seen[i.level + i.message] ? false : (seen[i.level + i.message] = true)));
}

export function blocking(issues: IIssue[]): boolean {
  return issues.some((i) => i.level === 'error');
}

/** Snap hours to the configured increment (e.g. 0.25) and clamp to 0–24. */
export function snapHours(value: number, increment: number): number {
  if (!isFinite(value) || value <= 0) return 0;
  const inc = increment > 0 ? increment : 0.25;
  return Math.min(24, Math.round(value / inc) * inc);
}

/** Parse what someone typed in an hours cell: "7.5", "7,5", "7:30", "7h30", "45m". */
export function parseHours(text: string): number {
  const s = (text || '').trim().toLowerCase().replace(',', '.');
  if (!s) return 0;
  let m = /^(\d{1,2})\s*[:h]\s*(\d{1,2})\s*m?$/.exec(s);
  if (m) return Number(m[1]) + Number(m[2]) / 60;
  m = /^(\d+)\s*m$/.exec(s);
  if (m) return Number(m[1]) / 60;
  m = /^(\d+(?:\.\d+)?)\s*h?$/.exec(s);
  if (m) return Number(m[1]);
  return NaN;
}

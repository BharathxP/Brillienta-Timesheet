// Who is due a pending-timesheet reminder and which projects need an end-date notice. Pure (tests/workflow.test.cjs).

import { IProject, ITimesheet, ITimesheetConfig, ITsUser } from '../models/types';
import { addDays, dayOfWeek, daysBetween, weekStartOf } from './dates';

export interface IReminderPlan {
  reminders: Array<{ user: ITsUser; week: string; hours: number }>;
  endNotices: Array<{ project: IProject; notice: string; days: number }>;
}

/** Who is due a reminder and which projects need an end-date notice (pure). */
export function planReminders(input: {
  cfg: ITimesheetConfig;
  now: Date;
  users: ITsUser[];
  sheets: ITimesheet[];
  projects: IProject[];
  sent: Record<string, string>;
  remindersOn: boolean;
  endDatesOn: boolean;
}): IReminderPlan {
  const { cfg, now } = input;
  const today = `${now.getFullYear()}-${('0' + (now.getMonth() + 1)).slice(-2)}-${('0' + now.getDate()).slice(-2)}`;
  const thisWeek = weekStartOf(today, cfg.weekStart);
  const plan: IReminderPlan = { reminders: [], endNotices: [] };
  const r = cfg.reminders;
  if (input.remindersOn && r.enabled) {
    const weeks: string[] = [];
    // Past weeks: once the reminder day and hour have come in the current week.
    const dayReached = daysBetween(thisWeek, today) >= (r.day - cfg.weekStart + 7) % 7 && now.getHours() >= r.hour;
    for (let i = 1; i <= Math.max(1, r.lookbackWeeks); i++) {
      if (i === 1 && !dayReached) continue;
      weeks.push(addDays(thisWeek, -7 * i));
    }
    // The current week, on its last working day.
    if (r.currentWeek) {
      const lastWork = [0, 1, 2, 3, 4, 5, 6].map((i) => addDays(thisWeek, i)).filter((d) => cfg.workDays.indexOf(dayOfWeek(d)) >= 0).pop();
      if (lastWork === today && now.getHours() >= r.hour) weeks.push(thisWeek);
    }
    input.users
      .filter((u) => u.active)
      .forEach((u) => {
        weeks.forEach((w) => {
          // Not for weeks before the person was added.
          if (u.created && u.created.substring(0, 10) > addDays(w, 6)) return;
          const s = input.sheets.filter((x) => x.userEmail.toLowerCase() === u.email.toLowerCase() && x.weekStart === w)[0];
          if (s && (s.status === 'Submitted' || s.status === 'Approved')) return;
          if (input.sent[`${u.email.toLowerCase()}|${w}`]) return;
          plan.reminders.push({ user: u, week: w, hours: s ? s.totalHours : 0 });
        });
      });
  }
  if (input.endDatesOn) {
    input.projects
      .filter((p) => p.status === 'Active' && p.endDate)
      .forEach((p) => {
        const left = daysBetween(today, p.endDate);
        if (left <= 0) {
          if (p.endNotices.indexOf('ended') < 0) plan.endNotices.push({ project: p, notice: 'ended', days: 0 });
          return;
        }
        // Only the nearest threshold counts: once the 7-day notice is out, the 14-day one is never sent late.
        const due = cfg.projectEndNotice.filter((n) => left <= n).sort((a, b) => a - b)[0];
        if (due !== undefined && p.endNotices.indexOf(String(due)) < 0) plan.endNotices.push({ project: p, notice: String(due), days: left });
      });
  }
  return plan;
}

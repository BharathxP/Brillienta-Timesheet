// Pending weekly timesheet reminders and project end-date notifications.
//
// They run in the browser of an administrator who opens the app (at most once an hour, one browser at a time), so no
// server or scheduled job is needed. Each reminder goes out once per person and week (ReminderLog setting); each end
// notice once per project and threshold (the project's EndNotices). The pure planning step is unit tested.

import { IProject, ITimesheet, ITimesheetConfig, ITsUser } from '../models/types';
import { SETTING_KEYS, SettingsStore } from './SettingsStore';
import { TimesheetService } from './TimesheetService';
import { Notifier } from './Notifier';
import { addDays, formatDay, formatWeek, weekStartOf } from './dates';
import { hours } from './format';

import { IReminderPlan, planReminders } from './reminderPlan';

export type { IReminderPlan };
export { planReminders };

export class ReminderRunner {
  private readonly _svc: TimesheetService;
  private readonly _settings: SettingsStore;
  private readonly _notifier: Notifier;

  public constructor(svc: TimesheetService, settings: SettingsStore, notifier: Notifier) {
    this._svc = svc;
    this._settings = settings;
    this._notifier = notifier;
  }

  /** Run if due. Returns what was sent (for the admin's toast and Admin › Notifications). */
  public async run(o: { cfg: ITimesheetConfig; users: ITsUser[]; projects: IProject[]; actor: string; remindersOn: boolean; endDatesOn: boolean; force?: boolean }): Promise<{ reminders: number; notices: number } | undefined> {
    const last = (await this._settings.fresh(SETTING_KEYS.reminderRun)) || '';
    const lastAt = Date.parse(last.split('|')[0]);
    if (!o.force && !isNaN(lastAt) && Date.now() - lastAt < 3600000) return undefined;
    await this._settings.set(SETTING_KEYS.reminderRun, `${new Date().toISOString()}|${o.actor}`);
    const now = new Date();
    const thisWeek = weekStartOf(`${now.getFullYear()}-${('0' + (now.getMonth() + 1)).slice(-2)}-${('0' + now.getDate()).slice(-2)}`, o.cfg.weekStart);
    const from = addDays(thisWeek, -7 * Math.max(1, o.cfg.reminders.lookbackWeeks));
    const { sheets } = o.remindersOn ? await this._svc.weeks(from, thisWeek) : { sheets: [] as ITimesheet[] };
    const sent = this._settings.getJson<Record<string, string>>(SETTING_KEYS.reminderLog) || {};
    const plan = planReminders({ cfg: o.cfg, now, users: o.users, sheets, projects: o.projects, sent, remindersOn: o.remindersOn, endDatesOn: o.endDatesOn });
    const byEmail = (e: string): ITsUser | undefined => o.users.filter((u) => u.email.toLowerCase() === e.toLowerCase())[0];
    for (const r of plan.reminders) {
      const manager = o.cfg.reminders.ccManager && r.user.managerEmail ? [r.user.managerEmail] : [];
      await this._notifier.send({ template: 'reminder', to: [{ email: r.user.email, name: r.user.name }], cc: manager, vars: { week: formatWeek(r.week), hours: hours(r.hours) }, path: `timesheet?week=${r.week}`, cta: 'Complete the timesheet' });
      sent[`${r.user.email.toLowerCase()}|${r.week}`] = now.toISOString();
      this._notifier.event('reminder.sent', { person: r.user.email, weekStart: r.week });
    }
    if (plan.reminders.length) {
      // Keep only the weeks that can still be chased.
      Object.keys(sent).forEach((k) => {
        if ((k.split('|')[1] || '') < addDays(from, -7)) delete sent[k];
      });
      await this._settings.setJson(SETTING_KEYS.reminderLog, sent);
    }
    let usedHours: Record<number, number> | undefined;
    for (const n of plan.endNotices) {
      const p = n.project;
      if (!usedHours) usedHours = await this._svc.hoursByProject().catch(() => ({}) as Record<number, number>);
      const to = p.managerEmails.map((e) => ({ email: e, name: (byEmail(e) || { name: e }).name }));
      const vars = { project: p.title, days: n.days, endDate: formatDay(p.endDate, true), used: hours(usedHours[p.id] || 0), planned: p.plannedHours ? hours(p.plannedHours) : '—' };
      await this._notifier.send({ template: n.notice === 'ended' ? 'projectEnded' : 'projectEnding', to, vars, path: `projects/${p.id}`, cta: 'Open the project' });
      // Smaller thresholds already passed are marked too, so a project created a day before its end gets one notice.
      const passed = o.cfg.projectEndNotice.filter((x) => n.notice === 'ended' || x >= Number(n.notice)).map(String);
      const notices = p.endNotices.concat(passed, n.notice === 'ended' ? ['ended'] : []).filter((v, i, a) => a.indexOf(v) === i);
      await this._svc.saveProject({ id: p.id, endNotices: notices });
      p.endNotices = notices;
      this._notifier.event('project.ending', { projectId: p.id, project: p.title, endDate: p.endDate, daysLeft: n.days });
    }
    return { reminders: plan.reminders.length, notices: plan.endNotices.length };
  }
}

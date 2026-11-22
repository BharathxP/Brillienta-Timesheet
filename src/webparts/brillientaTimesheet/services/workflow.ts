// The timesheet life cycle: save a draft, submit (rules, auto-approval, approvers told), approve / reject (bulk too),
// reopen, and clone a previous week. Used by the screens through the app context.

import { ISheetComment, ITimeEntry, ITimesheet, ITimesheetConfig, ITsUser, Role } from '../models/types';
import { TimesheetService } from './TimesheetService';
import { Notifier } from './Notifier';
import { IDirectory, approversFor, billRate, canApprove, costRate, sameEmail, shouldAutoApprove } from './access';
import { IIssue, blocking, validateWeek } from './rules';
import { entryTotal } from './mappers';
import { addDays, daysBetween, formatWeek, weekDays } from './dates';
import { hours } from './format';
import { TaskManagerService } from './TaskManagerService';

export interface IActor {
  email: string;
  name: string;
  role: Role;
}

export interface IWorkflowDeps {
  svc: TimesheetService;
  notifier: Notifier;
  config: () => ITimesheetConfig;
  dir: () => IDirectory;
  admins: () => string[];
  /** Plan features in force. */
  has: (feature: 'userRates' | 'autoApproval' | 'plannedHours' | 'taskManager') => boolean;
  /** Brillienta Task Manager (hours are written back there when configured). */
  taskManager?: TaskManagerService;
}

export interface ISubmitResult {
  sheet: ITimesheet;
  entries: ITimeEntry[];
  issues: IIssue[];
  autoApproved: boolean;
  approvers: string[];
}

const sum = (entries: ITimeEntry[], billableOnly: boolean): number => Math.round(entries.filter((e) => !billableOnly || e.billable).reduce((s, e) => s + entryTotal(e), 0) * 100) / 100;

export class Workflow {
  private readonly _d: IWorkflowDeps;

  public constructor(deps: IWorkflowDeps) {
    this._d = deps;
  }

  private _user(email: string): ITsUser | undefined {
    return this._d.dir().users.filter((u) => sameEmail(u.email, email))[0];
  }

  /** Save the week as it is (draft, or a rejected week being corrected). */
  public async save(owner: { email: string; name: string }, weekStart: string, sheet: ITimesheet | undefined, lines: ITimeEntry[], previous: ITimeEntry[]): Promise<{ sheet: ITimesheet; entries: ITimeEntry[] }> {
    const clean = lines
      .filter((l) => l.projectId || entryTotal(l) > 0)
      .map((l) => ({ ...l, userEmail: owner.email, weekStart, total: entryTotal(l) }));
    const base: ITimesheet = sheet || {
      id: 0,
      userEmail: owner.email,
      userName: owner.name,
      weekStart,
      status: 'Draft',
      totalHours: 0,
      billableHours: 0,
      submittedAt: '',
      decidedBy: '',
      decidedAt: '',
      autoApproved: false,
      comments: []
    };
    const next: ITimesheet = { ...base, userName: owner.name, totalHours: sum(clean, false), billableHours: sum(clean, true) };
    const id = await this._d.svc.saveSheet(next);
    const entries = await this._d.svc.saveEntries(id, clean, previous);
    return { sheet: { ...next, id }, entries };
  }

  /** Validate a week for submitting (the same checks the screen shows). */
  public check(owner: ITsUser | undefined, weekStart: string, lines: ITimeEntry[], timeOff: import('../models/types').ITimeOff[], timeOffActivityId: number): IIssue[] {
    const cfg = this._d.config();
    return validateWeek({
      cfg,
      user: owner,
      days: weekDays(weekStart),
      entries: lines.map((l) => ({ ...l, total: entryTotal(l) })),
      timeOff,
      timeOffActivityId,
      forSubmit: true,
      requireActivity: cfg.requireActivity,
      requiredColumns: cfg.customColumns.filter((c) => c.level === 'entry' && c.required).map((c) => ({ key: c.key, label: c.label }))
    });
  }

  /**
   * Submit: save, check the rules (errors stop it when enforced), capture rates, then either auto-approve (role-based)
   * or ask the approvers.
   */
  public async submit(actor: IActor, weekStart: string, sheet: ITimesheet | undefined, lines: ITimeEntry[], previous: ITimeEntry[], issues: IIssue[], comment: string): Promise<ISubmitResult> {
    if (blocking(issues)) throw new Error(issues.filter((i) => i.level === 'error')[0].message);
    const cfg = this._d.config();
    const saved = await this.save(actor, weekStart, sheet, lines, previous);
    const dir = this._d.dir();
    const owner = this._user(actor.email);
    // Rates are captured now, so later rate changes do not rewrite approved history.
    const stamped = saved.entries.map((e) => {
      const project = dir.projects.filter((p) => p.id === e.projectId)[0];
      return { ...e, costRate: costRate(cfg.costMode, owner, project, cfg.defaultCostRate), billRate: billRate(owner, project, cfg.defaultBillRate, this._d.has('userRates')) };
    });
    await this._d.svc.stampRates(stamped);
    const auto = this._d.has('autoApproval') && shouldAutoApprove(cfg, actor.role, saved.sheet.totalHours, issues);
    const now = new Date().toISOString();
    const comments: ISheetComment[] = saved.sheet.comments.concat(
      comment ? [{ at: now, by: actor.name, byEmail: actor.email, text: comment, action: 'submitted' }] : [{ at: now, by: actor.name, byEmail: actor.email, text: '', action: 'submitted' }]
    );
    if (auto) comments.push({ at: now, by: 'Auto-approval', byEmail: '', text: `Approved automatically (${actor.role} timesheets are auto-approved).`, action: 'auto-approved' });
    const next: ITimesheet = {
      ...saved.sheet,
      status: auto ? 'Approved' : 'Submitted',
      submittedAt: now,
      decidedBy: auto ? 'Auto-approval' : '',
      decidedAt: auto ? now : '',
      autoApproved: auto,
      comments
    };
    await this._d.svc.saveSheet(next);
    const approvers = auto ? [] : approversFor(actor.email, stamped, dir, cfg, this._d.admins());
    const vars = { person: actor.name, week: formatWeek(weekStart), hours: hours(next.totalHours), billable: hours(next.billableHours) };
    if (auto) {
      await this._d.notifier.send({ template: 'sheetApproved', to: [{ email: actor.email, name: actor.name }], vars: { ...vars, approver: 'Auto-approval', comment: '' }, path: `timesheet?week=${weekStart}`, cta: 'Open the timesheet' });
    } else {
      const to = approvers.map((e) => ({ email: e, name: (this._user(e) || { name: e }).name }));
      await this._d.notifier.send({ template: 'sheetSubmitted', to, vars, path: `approvals?sheet=${next.id}`, cta: 'Review the timesheet' });
    }
    this._d.notifier.event(auto ? 'timesheet.approved' : 'timesheet.submitted', { sheetId: next.id, person: actor.email, weekStart, hours: next.totalHours, billableHours: next.billableHours, autoApproved: auto });
    await this._d.svc.log(auto ? 'Timesheet auto-approved' : 'Timesheet submitted', `${actor.email} · ${weekStart} · ${next.totalHours} h`);
    await this._afterHours(stamped);
    const tm = this._d.config().taskManager.pushTime;
    const synced = tm === 'submitted' || (auto && tm === 'approved') ? await this._pushToTaskManager(stamped) : stamped;
    return { sheet: next, entries: synced, issues, autoApproved: auto, approvers };
  }

  /** Approve or reject sheets (one or many). Each must be one the actor may approve. */
  public async decide(actor: IActor, items: Array<{ sheet: ITimesheet; entries: ITimeEntry[] }>, approve: boolean, comment: string): Promise<number> {
    const cfg = this._d.config();
    const dir = this._d.dir();
    const now = new Date().toISOString();
    const allowed = items.filter((i) => i.sheet.status === 'Submitted' && canApprove(actor.role, actor.email, i.sheet, i.entries, dir, cfg));
    if (!approve && !comment.trim()) throw new Error('Say what needs to change when sending a timesheet back.');
    await this._d.svc.decideSheets(
      allowed.map((i) => ({
        id: i.sheet.id,
        row: {
          status: approve ? 'Approved' : 'Rejected',
          decidedBy: actor.name,
          decidedAt: now,
          comments: i.sheet.comments.concat([{ at: now, by: actor.name, byEmail: actor.email, text: comment, action: approve ? 'approved' : 'rejected' }])
        }
      }))
    );
    for (const i of allowed) {
      const vars = { approver: actor.name, week: formatWeek(i.sheet.weekStart), hours: hours(i.sheet.totalHours), comment };
      await this._d.notifier.send({
        template: approve ? 'sheetApproved' : 'sheetRejected',
        to: [{ email: i.sheet.userEmail, name: i.sheet.userName }],
        vars,
        path: `timesheet?week=${i.sheet.weekStart}`,
        cta: approve ? 'Open the timesheet' : 'Correct the timesheet'
      });
      this._d.notifier.event(approve ? 'timesheet.approved' : 'timesheet.rejected', { sheetId: i.sheet.id, person: i.sheet.userEmail, weekStart: i.sheet.weekStart, by: actor.email, comment });
    }
    await this._d.svc.log(approve ? 'Timesheets approved' : 'Timesheets rejected', `${allowed.length} by ${actor.email}`);
    if (approve && this._d.config().taskManager.pushTime === 'approved') {
      for (const i of allowed) await this._pushToTaskManager(i.entries);
    }
    return allowed.length;
  }

  /** Back to draft: the owner recalls a submitted week, or an approver / admin reopens an approved one. */
  public async reopen(actor: IActor, sheet: ITimesheet, reason: string, entries?: ITimeEntry[]): Promise<ITimesheet> {
    const now = new Date().toISOString();
    // Hours written to Task Manager are taken back; they go again when the week is approved (or submitted) again.
    if (entries && entries.some((e) => e.tmSync && e.tmSync.pushed && e.tmSync.pushed.length)) {
      await this._pushToTaskManager(entries.map((e) => ({ ...e, hours: {} })), true);
    }
    const next: ITimesheet = { ...sheet, status: 'Draft', decidedBy: '', decidedAt: '', autoApproved: false, comments: sheet.comments.concat([{ at: now, by: actor.name, byEmail: actor.email, text: reason, action: 'reopened' }]) };
    await this._d.svc.saveSheet(next);
    await this._d.svc.log('Timesheet reopened', `${sheet.userEmail} · ${sheet.weekStart} by ${actor.email}`);
    return next;
  }

  /**
   * Clone a week: the lines of `source` moved onto the target week's days (same weekday), keeping project, task,
   * activity, billable and custom columns; hours optional; lines on projects that are no longer open are dropped.
   */
  public cloneLines(source: ITimeEntry[], sourceWeek: string, targetWeek: string, withHours: boolean, openProjectIds: number[]): ITimeEntry[] {
    const shift = daysBetween(sourceWeek, targetWeek);
    return source
      .filter((e) => openProjectIds.indexOf(e.projectId) >= 0)
      .map((e) => {
        const h: Record<string, number> = {};
        if (withHours) Object.keys(e.hours).forEach((d) => (h[addDays(d, shift)] = e.hours[d]));
        return { ...e, id: 0, sheetId: 0, weekStart: targetWeek, hours: h, notes: {}, total: 0, costRate: 0, billRate: 0, tmSync: {} };
      });
  }

  /**
   * Write lines' hours to Brillienta Task Manager (see TaskManagerService.pushWeek) and remember what was sent.
   * Problems are reported, never thrown: the timesheet action has already happened.
   */
  private async _pushToTaskManager(entries: ITimeEntry[], withdraw: boolean = false): Promise<ITimeEntry[]> {
    const tm = this._d.taskManager;
    if (!tm || !tm.enabled || !this._d.has('taskManager')) return entries;
    try {
      const r = await tm.pushWeek(entries, this._d.dir());
      if (r.changed.length) await this._d.svc.saveTmSync(r.changed.map((c) => (withdraw ? { ...c, tmSync: {} } : c)));
      if (r.problems.length) this._d.notifier.onProblem(`Hours not fully written to Task Manager: ${r.problems[0]}`);
      return entries.map((e) => {
        const c = r.changed.filter((x) => x.id === e.id)[0];
        return c ? { ...e, tmSync: withdraw ? {} : c.tmSync } : e;
      });
    } catch (e) {
      this._d.notifier.onProblem('Hours not written to Task Manager: ' + (e as Error).message);
      return entries;
    }
  }

  /** Planned-hours alerts: tell project managers once when logged hours pass the plan. */
  private async _afterHours(entries: ITimeEntry[]): Promise<void> {
    const cfg = this._d.config();
    if (!cfg.plannedHoursAlerts || !this._d.has('plannedHours')) return;
    const ids = entries.map((e) => e.projectId).filter((v, i, a) => v && a.indexOf(v) === i);
    const dir = this._d.dir();
    const planned = dir.projects.filter((p) => ids.indexOf(p.id) >= 0 && p.plannedHours > 0 && p.endNotices.indexOf('planned') < 0);
    if (!planned.length) return;
    try {
      const used = await this._d.svc.hoursByProject();
      for (const p of planned) {
        if ((used[p.id] || 0) <= p.plannedHours) continue;
        await this._d.notifier.send({
          template: 'plannedHoursExceeded',
          to: p.managerEmails.map((e) => ({ email: e, name: (this._user(e) || { name: e }).name })),
          vars: { project: p.title, used: hours(used[p.id]), planned: hours(p.plannedHours) },
          path: `projects/${p.id}`,
          cta: 'Open the project'
        });
        await this._d.svc.saveProject({ id: p.id, endNotices: p.endNotices.concat('planned') });
        p.endNotices = p.endNotices.concat('planned');
      }
    } catch {
      // alerts never block a submit
    }
  }
}

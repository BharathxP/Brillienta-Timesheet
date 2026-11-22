// Brillienta Task Manager integration — the part that reads and writes Task Manager's lists. The mapping and
// planning rules are in taskManagerSync.ts.

import { ITaskManagerConfig, ITimeEntry } from '../models/types';
import { SPClient, odataString } from './SPClient';
import { IReference, TimesheetService } from './TimesheetService';
import { ITmProject, ITmTask, ITmTimeEntry, importLines, planProjects, planTasks, pushItems, toTmProject, toTmTask, toTmTimeEntry } from './taskManagerSync';
import { PALETTE } from '../schema/defaults';

export interface ISyncResult {
  projectsCreated: number;
  projectsUpdated: number;
  tasksCreated: number;
  tasksUpdated: number;
}

const MAX_IN = 40;

export class TaskManagerService {
  private readonly _sp: SPClient;
  private readonly _cfg: () => ITaskManagerConfig;

  public constructor(sp: SPClient, cfg: () => ITaskManagerConfig) {
    this._sp = sp;
    this._cfg = cfg;
  }

  private _client(): SPClient {
    const c = this._cfg();
    return c.siteUrl ? this._sp.forWeb(c.siteUrl) : this._sp;
  }

  private _list(suffix: string): string {
    const c = this._cfg();
    return this._client().list((c.prefix || 'TM_') + suffix);
  }

  public get enabled(): boolean {
    return this._cfg().enabled;
  }

  /** Check the Task Manager lists can be read (and written, for time entries). Returns problems; empty = fine. */
  public async test(): Promise<string[]> {
    const problems: string[] = [];
    const sp = this._client();
    for (const suffix of ['Projects', 'Tasks', 'TimeEntries']) {
      try {
        const info = await sp.get<{ ItemCount: number; EffectiveBasePermissions?: { Low: string } }>(`${this._list(suffix)}?$select=ItemCount,EffectiveBasePermissions`);
        // AddListItems = 0x2: needed on TimeEntries to write hours back.
        if (suffix === 'TimeEntries' && info.EffectiveBasePermissions && !(Number(info.EffectiveBasePermissions.Low) & 0x2)) {
          problems.push(`You can read ${this._cfg().prefix}TimeEntries but not add to it, so hours cannot be written back from your account.`);
        }
      } catch (e) {
        problems.push(`${this._cfg().prefix}${suffix} could not be read: ${(e as Error).message}`);
      }
    }
    return problems;
  }

  public async projects(): Promise<ITmProject[]> {
    const rows = await this._client().getAll<Record<string, unknown>>(
      `${this._list('Projects')}/items?$select=Id,Title,ProjectCode,ProjectStatus,PlannedStart,PlannedEnd,BudgetHours,Details,ColorHex,ProjectOwner/EMail,ProjectMembers/EMail&$expand=ProjectOwner,ProjectMembers&$top=2000`,
      10000
    );
    return rows.map(toTmProject);
  }

  public async tasks(tmProjectIds: number[]): Promise<ITmTask[]> {
    const out: ITmTask[] = [];
    for (let i = 0; i < tmProjectIds.length; i += MAX_IN) {
      const f = tmProjectIds
        .slice(i, i + MAX_IN)
        .map((id) => `ProjectId eq ${id}`)
        .join(' or ');
      const rows = await this._client().getAll<Record<string, unknown>>(
        `${this._list('Tasks')}/items?$select=Id,Title,TaskNumber,ProjectId,StatusKind,AssigneeEmail,DueDate,EstimateHours,SpentMinutes&$filter=${encodeURIComponent(f)}&$top=5000`,
        50000
      );
      out.push(...rows.map(toTmTask));
    }
    return out;
  }

  /** Time a person logged in Task Manager between two days. */
  public async timeOf(email: string, from: string, to: string): Promise<ITmTimeEntry[]> {
    const f = `WorkDate ge ${odataString(from)} and WorkDate le ${odataString(to)} and Author/EMail eq ${odataString(email)}`;
    const rows = await this._client().getAll<Record<string, unknown>>(
      `${this._list('TimeEntries')}/items?$select=Id,TaskId,ProjectId,DurationMinutes,WorkNote,Billable,WorkDate,Author/EMail&$expand=Author&$filter=${encodeURIComponent(f)}&$top=2000`,
      5000
    );
    return rows.map(toTmTimeEntry);
  }

  /** Bring Task Manager projects and tasks into the timesheet (as the settings say). */
  public async sync(svc: TimesheetService, ref: IReference): Promise<ISyncResult> {
    const c = this._cfg();
    const result: ISyncResult = { projectsCreated: 0, projectsUpdated: 0, tasksCreated: 0, tasksUpdated: 0 };
    let projects = ref.projects;
    if (c.syncProjects) {
      const plan = planProjects(await this.projects(), ref.projects, c.projectStatuses, PALETTE);
      if (plan.create.length) await svc.addProjects(plan.create);
      for (const u of plan.update) await svc.saveProject({ id: u.id, ...u.patch });
      result.projectsCreated = plan.create.length;
      result.projectsUpdated = plan.update.length;
      if (plan.create.length || plan.update.length) projects = (await svc.loadReference()).projects;
    }
    if (c.syncTasks) {
      const linked = projects.filter((p) => p.tmProjectId > 0);
      if (linked.length) {
        const tmTasks = await this.tasks(linked.map((p) => p.tmProjectId));
        const plan = planTasks(tmTasks, ref.tasks, projects);
        if (plan.create.length) await svc.addTasks(plan.create);
        for (const u of plan.update) await svc.saveTask({ id: u.id, ...u.patch });
        result.tasksCreated = plan.create.length;
        result.tasksUpdated = plan.update.length;
      }
    }
    await svc.log('Task Manager sync', `${result.projectsCreated} projects added, ${result.projectsUpdated} updated; ${result.tasksCreated} tasks added, ${result.tasksUpdated} updated`);
    return result;
  }

  /** Lines for a person's week from the time they logged in Task Manager. */
  public async importWeek(email: string, weekStart: string, weekEnd: string, ref: IReference, current: ITimeEntry[]): Promise<{ lines: ITimeEntry[]; skipped: number }> {
    const done = current.reduce<number[]>((a, l) => a.concat((l.tmSync && l.tmSync.imported) || []), []);
    return importLines(await this.timeOf(email, weekStart, weekEnd), weekStart, email, ref.tasks, ref.projects, done);
  }

  /**
   * Write a week's hours to Task Manager: for each line on a linked task, remove what it sent before (entries and
   * minutes), then add one time entry per day and add the minutes to the task's time spent. Returns the lines whose
   * sync state changed (to save on the timesheet lines) and the problems met.
   */
  public async pushWeek(lines: ITimeEntry[], ref: Pick<IReference, 'projects' | 'tasks'>): Promise<{ changed: ITimeEntry[]; problems: string[] }> {
    const sp = this._client();
    const changed: ITimeEntry[] = [];
    const problems: string[] = [];
    for (const line of lines) {
      const { tmTaskId, items } = pushItems(line, ref.tasks);
      const before = line.tmSync || {};
      if (!tmTaskId && !(before.pushed && before.pushed.length)) continue;
      try {
        // 1. take back what this line sent before (re-approval after a reopen, or the task link changed)
        if (before.pushed && before.pushed.length) {
          for (const id of before.pushed) await sp.del(`${this._list('TimeEntries')}/items(${id})`).catch(() => undefined);
          if (before.pushedTaskId && before.pushedMinutes) await this._addSpent(before.pushedTaskId, -before.pushedMinutes);
        }
        // 2. send the line as it is now
        const pushed: number[] = [];
        let minutes = 0;
        if (tmTaskId) {
          const project = ref.projects.filter((p) => p.id === line.projectId)[0];
          for (const it of items) {
            const r = await sp.post<{ Id: number }>(`${this._list('TimeEntries')}/items`, {
              Title: 'Brillienta Timesheet',
              TaskId: tmTaskId,
              ProjectId: project ? project.tmProjectId : 0,
              DurationMinutes: it.minutes,
              WorkNote: (it.note ? it.note + ' — ' : '') + `from Brillienta Timesheet (${line.userEmail})`,
              Billable: line.billable,
              WorkDate: it.workDate
            });
            pushed.push(r.Id);
            minutes += it.minutes;
          }
          if (minutes) await this._addSpent(tmTaskId, minutes);
        }
        changed.push({ ...line, tmSync: { pushed, pushedMinutes: minutes, pushedTaskId: tmTaskId } });
      } catch (e) {
        problems.push((e as Error).message);
      }
    }
    return { changed, problems };
  }

  private async _addSpent(tmTaskId: number, minutes: number): Promise<void> {
    const path = `${this._list('Tasks')}/items(${tmTaskId})`;
    const t = await this._client().get<{ SpentMinutes: number }>(`${path}?$select=SpentMinutes`);
    await this._client().merge(path, { SpentMinutes: Math.max(0, (Number(t.SpentMinutes) || 0) + minutes), LastActivityAt: new Date().toISOString() });
  }
}

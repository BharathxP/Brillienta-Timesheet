// Every read and write of the timesheet lists. Screens never build REST URLs themselves.

import { SPClient, odataString, IBatchWrite } from './SPClient';
import { ListKey, listTitle, selectFor } from '../schema/schema';
import {
  IActivity,
  IAuditEntry,
  INotification,
  IProgram,
  IProject,
  ISavedReport,
  ITask,
  ITicket,
  ITimeEntry,
  ITimesheet,
  ITsUser
} from '../models/types';
import * as M from './mappers';

export interface IReference {
  users: ITsUser[];
  programs: IProgram[];
  projects: IProject[];
  tasks: ITask[];
  activities: IActivity[];
}

const MAX_IN = 40;

export class TimesheetService {
  public readonly sp: SPClient;
  public readonly prefix: string;
  /** Name written to the audit log. */
  public actor = '';

  public constructor(sp: SPClient, prefix: string) {
    this.sp = sp;
    this.prefix = prefix;
  }

  public title(key: ListKey): string {
    return listTitle(this.prefix, key);
  }

  public items(key: ListKey): string {
    return `${this.sp.list(this.title(key))}/items`;
  }

  private _query(key: ListKey, filter?: string, top: number = 5000, orderBy: string = 'Id'): string {
    return `${this.items(key)}?$select=${selectFor(key)}&$expand=Author&$top=${top}&$orderby=${orderBy}` + (filter ? `&$filter=${encodeURIComponent(filter)}` : '');
  }

  private async _all(key: ListKey, filter?: string, max: number = 50000): Promise<M.Row[]> {
    return this.sp.getAll<M.Row>(this._query(key, filter), max);
  }

  private async _add(key: ListKey, row: M.Row): Promise<number> {
    const r = await this.sp.post<{ Id: number }>(this.items(key), row);
    return r.Id;
  }

  private async _update(key: ListKey, id: number, row: M.Row): Promise<void> {
    await this.sp.merge(`${this.items(key)}(${id})`, row);
  }

  private async _delete(key: ListKey, id: number): Promise<void> {
    await this.sp.del(`${this.items(key)}(${id})`);
  }

  /** Writes many rows in $batch requests; throws with the first failure after all were attempted. */
  private async _batch(writes: IBatchWrite[]): Promise<number[]> {
    if (!writes.length) return [];
    const results = await this.sp.batch(writes);
    const failed = results.filter((r) => r.status >= 400 || r.status === 0);
    if (failed.length) throw new Error(`${failed.length} of ${writes.length} changes failed: ${failed[0].error || failed[0].status}`);
    return results.map((r) => {
      const body = r.data as { Id?: number } | undefined;
      return body && body.Id ? body.Id : 0;
    });
  }

  // ------------------------------------------------------------------ reference data

  public async loadReference(): Promise<IReference> {
    const [users, programs, projects, tasks, activities] = await Promise.all([
      this._all('users').then((r) => r.map(M.toUser)),
      this._all('programs').then((r) => r.map(M.toProgram)),
      this._all('projects').then((r) => r.map(M.toProject)),
      this._all('tasks').then((r) => r.map(M.toTask)),
      this._all('activities').then((r) => r.map(M.toActivity))
    ]);
    const byName = (a: { title: string }, b: { title: string }): number => a.title.localeCompare(b.title);
    users.sort((a, b) => a.name.localeCompare(b.name));
    projects.sort(byName);
    programs.sort(byName);
    tasks.sort(byName);
    return { users, programs, projects, tasks, activities };
  }

  public async saveUser(u: Partial<ITsUser> & { id?: number }): Promise<number> {
    if (u.id) {
      await this._update('users', u.id, M.fromUser(u));
      return u.id;
    }
    return this._add('users', M.fromUser(u));
  }

  public async addUsers(users: Array<Partial<ITsUser>>): Promise<void> {
    await this._batch(users.map((u) => ({ method: 'POST', path: this.items('users'), body: M.fromUser(u) }) as IBatchWrite));
  }

  public async deleteUser(id: number): Promise<void> {
    await this._delete('users', id);
  }

  public async saveProgram(p: Partial<IProgram> & { id?: number }): Promise<number> {
    if (p.id) {
      await this._update('programs', p.id, M.fromProgram(p));
      return p.id;
    }
    return this._add('programs', M.fromProgram(p));
  }

  public async deleteProgram(id: number): Promise<void> {
    await this._delete('programs', id);
  }

  public async saveProject(p: Partial<IProject> & { id?: number }): Promise<number> {
    if (p.id) {
      await this._update('projects', p.id, M.fromProject(p));
      return p.id;
    }
    return this._add('projects', M.fromProject(p));
  }

  public async addProjects(list: Array<Partial<IProject>>): Promise<number[]> {
    return this._batch(list.map((p) => ({ method: 'POST', path: this.items('projects'), body: M.fromProject(p) }) as IBatchWrite));
  }

  /** Deletes a project with its tasks. Logged time stays (reports show it as "(deleted project)"). */
  public async deleteProject(id: number, taskIds: number[]): Promise<void> {
    await this._batch(taskIds.map((t) => ({ method: 'DELETE', path: `${this.items('tasks')}(${t})` }) as IBatchWrite));
    await this._delete('projects', id);
  }

  public async saveTask(t: Partial<ITask> & { id?: number }): Promise<number> {
    if (t.id) {
      await this._update('tasks', t.id, M.fromTask(t));
      return t.id;
    }
    return this._add('tasks', M.fromTask(t));
  }

  public async addTasks(list: Array<Partial<ITask>>): Promise<number[]> {
    return this._batch(list.map((t) => ({ method: 'POST', path: this.items('tasks'), body: M.fromTask(t) }) as IBatchWrite));
  }

  public async deleteTask(id: number): Promise<void> {
    await this._delete('tasks', id);
  }

  public async saveActivity(a: Partial<IActivity> & { id?: number }): Promise<number> {
    if (a.id) {
      await this._update('activities', a.id, M.fromActivity(a));
      return a.id;
    }
    return this._add('activities', M.fromActivity(a));
  }

  public async addActivities(list: Array<Partial<IActivity>>): Promise<void> {
    await this._batch(list.map((a) => ({ method: 'POST', path: this.items('activities'), body: M.fromActivity(a) }) as IBatchWrite));
  }

  public async deleteActivity(id: number): Promise<void> {
    await this._delete('activities', id);
  }

  // ------------------------------------------------------------------ timesheets

  public async getSheet(email: string, weekStart: string): Promise<{ sheet?: ITimesheet; entries: ITimeEntry[] }> {
    const f = `UserEmail eq ${odataString(email)} and WeekStart eq ${odataString(weekStart)}`;
    const [sheets, entries] = await Promise.all([this._all('timesheets', f, 10), this._all('entries', f, 500)]);
    return { sheet: sheets.map(M.toSheet)[0], entries: entries.map(M.toEntry) };
  }

  /** Sheets of a person (most recent first). */
  public async sheetsOf(email: string, top: number = 60): Promise<ITimesheet[]> {
    const rows = await this.sp.get<{ value: M.Row[] }>(this._query('timesheets', `UserEmail eq ${odataString(email)}`, top, 'WeekStart desc'));
    return rows.value.map(M.toSheet);
  }

  /** Sheets with a status (approvals queue). */
  public async sheetsByStatus(status: ITimesheet['status']): Promise<ITimesheet[]> {
    return (await this._all('timesheets', `SheetStatus eq ${odataString(status)}`, 5000)).map(M.toSheet);
  }

  /** Sheets and their lines for weeks in [fromWeek, toWeek] (week-start keys sort as text). */
  public async weeks(fromWeek: string, toWeek: string, emails?: string[]): Promise<{ sheets: ITimesheet[]; entries: ITimeEntry[] }> {
    const range = `WeekStart ge ${odataString(fromWeek)} and WeekStart le ${odataString(toWeek)}`;
    const people = (emails || []).slice(0, MAX_IN);
    const who = people.length && emails && emails.length <= MAX_IN ? ` and (${people.map((e) => `UserEmail eq ${odataString(e)}`).join(' or ')})` : '';
    const [sheets, entries] = await Promise.all([this._all('timesheets', range + who), this._all('entries', range + who)]);
    let s = sheets.map(M.toSheet);
    let e = entries.map(M.toEntry);
    if (emails && emails.length > MAX_IN) {
      const set = emails.map((x) => x.toLowerCase());
      s = s.filter((x) => set.indexOf(x.userEmail.toLowerCase()) >= 0);
      e = e.filter((x) => set.indexOf(x.userEmail.toLowerCase()) >= 0);
    }
    return { sheets: s, entries: e };
  }

  public async entriesOfSheets(sheetIds: number[]): Promise<ITimeEntry[]> {
    const out: ITimeEntry[] = [];
    for (let i = 0; i < sheetIds.length; i += MAX_IN) {
      const chunk = sheetIds.slice(i, i + MAX_IN);
      const rows = await this._all('entries', chunk.map((id) => `SheetId eq ${id}`).join(' or '));
      out.push(...rows.map(M.toEntry));
    }
    return out;
  }

  /** Lines of one project (project report, planned hours). */
  public async entriesOfProject(projectId: number): Promise<{ sheets: ITimesheet[]; entries: ITimeEntry[] }> {
    const entries = (await this._all('entries', `ProjectId eq ${projectId}`)).map(M.toEntry);
    const ids = entries.map((e) => e.sheetId).filter((v, i, a) => a.indexOf(v) === i);
    const sheets: ITimesheet[] = [];
    for (let i = 0; i < ids.length; i += MAX_IN) {
      const rows = await this._all('timesheets', ids.slice(i, i + MAX_IN).map((id) => `Id eq ${id}`).join(' or '));
      sheets.push(...rows.map(M.toSheet));
    }
    return { sheets, entries };
  }

  /** Total logged hours per project (planned hours tracking). Reads only two columns. */
  public async hoursByProject(): Promise<Record<number, number>> {
    const rows = await this.sp.getAll<{ ProjectId: number; TotalHours: number }>(`${this.items('entries')}?$select=ProjectId,TotalHours&$top=5000`, 200000);
    const out: Record<number, number> = {};
    rows.forEach((r) => (out[r.ProjectId] = Math.round(((out[r.ProjectId] || 0) + (Number(r.TotalHours) || 0)) * 100) / 100));
    return out;
  }

  /** Hours per task (task progress). */
  public async hoursByTask(): Promise<Record<number, number>> {
    const rows = await this.sp.getAll<{ TaskId: number; TotalHours: number }>(`${this.items('entries')}?$select=TaskId,TotalHours&$filter=TaskId gt 0&$top=5000`, 200000);
    const out: Record<number, number> = {};
    rows.forEach((r) => (out[r.TaskId] = Math.round(((out[r.TaskId] || 0) + (Number(r.TotalHours) || 0)) * 100) / 100));
    return out;
  }

  /** Create or update the sheet row. */
  public async saveSheet(s: Partial<ITimesheet> & { id?: number; userEmail: string; weekStart: string }): Promise<number> {
    if (s.id) {
      await this._update('timesheets', s.id, M.fromSheet(s));
      return s.id;
    }
    return this._add('timesheets', M.fromSheet(s));
  }

  /**
   * Save a week's lines: update changed ones, add new ones (id 0), delete removed ones. One $batch for everything.
   * Returns the lines with their ids.
   */
  public async saveEntries(sheetId: number, lines: ITimeEntry[], previous: ITimeEntry[]): Promise<ITimeEntry[]> {
    const writes: IBatchWrite[] = [];
    const kept = lines.filter((l) => l.id);
    previous.forEach((p) => {
      if (!kept.some((k) => k.id === p.id)) writes.push({ method: 'DELETE', path: `${this.items('entries')}(${p.id})` });
    });
    const order: number[] = [];
    lines.forEach((l, i) => {
      const row = M.fromEntry({ ...l, sheetId });
      if (l.id) {
        const before = previous.filter((p) => p.id === l.id)[0];
        if (before && JSON.stringify(M.fromEntry({ ...before, sheetId })) === JSON.stringify(row)) return;
        writes.push({ method: 'MERGE', path: `${this.items('entries')}(${l.id})`, body: row });
      } else {
        order.push(i);
        writes.push({ method: 'POST', path: this.items('entries'), body: row });
      }
    });
    if (!writes.length) return lines.map((l) => ({ ...l, sheetId }));
    const results = await this.sp.batch(writes);
    const failed = results.filter((r) => r.status >= 400 || r.status === 0);
    if (failed.length) throw new Error(`Some lines could not be saved: ${failed[0].error || failed[0].status}`);
    const posts = results.filter((_, i) => writes[i].method === 'POST');
    const out = lines.map((l) => ({ ...l, sheetId }));
    order.forEach((lineIndex, j) => {
      const body = posts[j] && (posts[j].data as { Id?: number } | undefined);
      if (body && body.Id) out[lineIndex].id = body.Id;
    });
    return out;
  }

  /** Approve / reject many sheets in one go (bulk approvals). */
  public async decideSheets(updates: Array<{ id: number; row: Partial<ITimesheet> }>): Promise<void> {
    await this._batch(updates.map((u) => ({ method: 'MERGE', path: `${this.items('timesheets')}(${u.id})`, body: M.fromSheet(u.row) }) as IBatchWrite));
  }

  /** Remember what each line sent to (or imported from) Brillienta Task Manager. */
  public async saveTmSync(lines: Array<{ id: number; tmSync: unknown }>): Promise<void> {
    await this._batch(lines.filter((l) => l.id).map((l) => ({ method: 'MERGE', path: `${this.items('entries')}(${l.id})`, body: { TmSync: JSON.stringify(l.tmSync || {}) } }) as IBatchWrite));
  }

  /** Capture rates on the lines at submit time. */
  public async stampRates(entries: Array<{ id: number; costRate: number; billRate: number }>): Promise<void> {
    await this._batch(entries.map((e) => ({ method: 'MERGE', path: `${this.items('entries')}(${e.id})`, body: { CostRate: e.costRate, BillRate: e.billRate } }) as IBatchWrite));
  }

  public async deleteSheet(sheet: ITimesheet, entries: ITimeEntry[]): Promise<void> {
    await this._batch(entries.map((e) => ({ method: 'DELETE', path: `${this.items('entries')}(${e.id})` }) as IBatchWrite));
    await this._delete('timesheets', sheet.id);
  }

  // ------------------------------------------------------------------ support, notifications, audit, reports

  public async tickets(): Promise<ITicket[]> {
    const rows = await this.sp.get<{ value: M.Row[] }>(this._query('tickets', undefined, 500, 'Id desc'));
    return rows.value.map(M.toTicket);
  }

  public async addTicket(t: Pick<ITicket, 'title' | 'details' | 'category' | 'priority' | 'raisedByEmail'>): Promise<number> {
    return this._add('tickets', { Title: t.title, Details: t.details, Category: t.category, Priority: t.priority, RaisedByEmail: t.raisedByEmail, TicketStatus: 'Open' });
  }

  public async updateTicket(id: number, row: { TicketStatus?: string; Reply?: string }): Promise<void> {
    await this._update('tickets', id, row);
  }

  public async notifications(email: string, top: number = 40): Promise<INotification[]> {
    const rows = await this.sp.get<{ value: M.Row[] }>(this._query('notifications', `Recipient eq ${odataString(email)}`, top, 'Id desc'));
    return rows.value.map(M.toNotification);
  }

  public async notify(list: Array<{ recipient: string; title: string; body: string; link: string; kind: string }>): Promise<void> {
    await this._batch(list.map((n) => ({ method: 'POST', path: this.items('notifications'), body: { Title: n.title.substring(0, 255), Body: n.body, LinkPath: n.link, Recipient: n.recipient, Kind: n.kind, IsRead: false } }) as IBatchWrite));
  }

  public async markRead(ids: number[]): Promise<void> {
    await this._batch(ids.map((id) => ({ method: 'MERGE', path: `${this.items('notifications')}(${id})`, body: { IsRead: true } }) as IBatchWrite));
  }

  public async queueEmail(m: { to: string[]; cc: string[]; subject: string; html: string; kind: string }): Promise<void> {
    // The flow sends one email to every address in MailTo (copies included).
    await this._add('emailQueue', { Title: m.subject.substring(0, 255), MailTo: m.to.concat(m.cc).join(';'), MailBody: m.html, TemplateKey: m.kind, MailStatus: 'Queued' });
  }

  public async emailQueue(top: number = 50): Promise<Array<{ id: number; subject: string; to: string; status: string; created: string }>> {
    const rows = await this.sp.get<{ value: M.Row[] }>(this._query('emailQueue', undefined, top, 'Id desc'));
    return rows.value.map((r) => ({ id: Number(r.Id), subject: String(r.Title || ''), to: String(r.MailTo || ''), status: String(r.MailStatus || ''), created: String(r.Created || '') }));
  }

  public async log(action: string, details: string = ''): Promise<void> {
    try {
      await this._add('audit', { Title: action.substring(0, 255), Details: details, ByName: this.actor });
    } catch {
      // the audit log never blocks the change itself
    }
  }

  public async audit(top: number = 200): Promise<IAuditEntry[]> {
    const rows = await this.sp.get<{ value: M.Row[] }>(this._query('audit', undefined, top, 'Id desc'));
    return rows.value.map(M.toAudit);
  }

  public async reports(email: string): Promise<ISavedReport[]> {
    const rows = await this._all('reports', `OwnerEmail eq ${odataString(email)} or IsShared eq 1`, 500);
    return rows.map(M.toReport);
  }

  public async saveReport(r: Omit<ISavedReport, 'id'> & { id?: number }): Promise<number> {
    const row = { Title: r.title, OwnerEmail: r.owner, IsShared: r.shared, ReportKind: r.kind, Definition: JSON.stringify(r.definition) };
    if (r.id) {
      await this._update('reports', r.id, row);
      return r.id;
    }
    return this._add('reports', row);
  }

  public async deleteReport(id: number): Promise<void> {
    await this._delete('reports', id);
  }

  /** Absolute REST URL of a list's items (API page, Power BI). */
  public restUrl(key: ListKey): string {
    return this.sp.api(this.items(key));
  }
}

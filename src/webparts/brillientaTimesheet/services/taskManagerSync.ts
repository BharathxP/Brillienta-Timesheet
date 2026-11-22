// Brillienta Task Manager integration — the pure part (tests/taskmanager.test.cjs).
//
// Task Manager keeps its data in SharePoint lists prefixed TM_ (repository BharathxP/Brillienta-TaskManager):
//   TM_Projects     Title, ProjectCode, ProjectStatus (Planning/Active/OnHold/Completed/Cancelled), PlannedStart,
//                   PlannedEnd, BudgetHours, ProjectOwner, ProjectMembers, Details, ColorHex
//   TM_Tasks        Title, TaskNumber, ProjectId, StatusKind (NotStarted/InProgress/Waiting/Done/Cancelled),
//                   AssigneeEmail, DueDate, EstimateHours, SpentMinutes
//   TM_TimeEntries  TaskId, ProjectId, DurationMinutes, WorkNote, Billable, WorkDate (yyyy-mm-dd); Author = who logged it
//
// Direction of each piece of data:
//   projects and tasks   Task Manager → Timesheet (linked by TmProjectId / TmTaskId; Task Manager owns names, status,
//                        dates, people and estimates; Timesheet owns rates, billable and custom columns)
//   hours                Timesheet → Task Manager (one TM_TimeEntries row per line and day, and the task's SpentMinutes),
//                        when a week is approved (or submitted). Re-sending replaces what the line sent before.
//   logged time          Task Manager → a person's week, on request ("Import from Task Manager"); imported lines are
//                        never written back, so nothing is counted twice.

import { DateKey, IProject, ITask, ITimeEntry, ProjectStatus, TaskStatus } from '../models/types';
import { toKey, weekDays } from './dates';

export interface ITmProject {
  id: number;
  title: string;
  code: string;
  status: string;
  start: DateKey;
  end: DateKey;
  budgetHours: number;
  ownerEmail: string;
  memberEmails: string[];
  description: string;
  color: string;
}

export interface ITmTask {
  id: number;
  title: string;
  number: string;
  projectId: number;
  statusKind: string;
  assigneeEmail: string;
  due: DateKey;
  estimateHours: number;
  spentMinutes: number;
}

export interface ITmTimeEntry {
  id: number;
  taskId: number;
  projectId: number;
  minutes: number;
  note: string;
  billable: boolean;
  workDate: DateKey;
  authorEmail: string;
}

const lc = (s: string): string => (s || '').trim().toLowerCase();

/** The calendar day of a Task Manager date-time (stored as an instant). */
export function tmDay(iso: unknown): DateKey {
  const s = typeof iso === 'string' ? iso : '';
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  return isNaN(d.getTime()) ? '' : toKey(d);
}

export function taskStatusFromTm(kind: string): TaskStatus {
  switch (kind) {
    case 'InProgress':
    case 'Waiting':
      return 'In progress';
    case 'Done':
    case 'Cancelled':
      return 'Done';
    default:
      return 'Open';
  }
}

export function projectStatusFromTm(status: string): ProjectStatus {
  if (status === 'Completed' || status === 'Cancelled') return 'Completed';
  if (status === 'OnHold') return 'Disabled';
  return 'Active';
}

type Row = Record<string, unknown>;
const str = (v: unknown): string => (v === undefined || v === null ? '' : String(v));
const num = (v: unknown): number => (isFinite(Number(v)) ? Number(v) : 0);
const person = (v: unknown): string => (v && typeof v === 'object' ? str((v as Row).EMail || (v as Row).Email) : '');
const people = (v: unknown): string[] => {
  const list = v && typeof v === 'object' ? ((v as Row).results as unknown[]) || (Array.isArray(v) ? (v as unknown[]) : []) : [];
  return list.map(person).filter((e) => !!e);
};

export function toTmProject(r: Row): ITmProject {
  return {
    id: num(r.Id),
    title: str(r.Title),
    code: str(r.ProjectCode),
    status: str(r.ProjectStatus) || 'Active',
    start: tmDay(r.PlannedStart),
    end: tmDay(r.PlannedEnd),
    budgetHours: num(r.BudgetHours),
    ownerEmail: person(r.ProjectOwner),
    memberEmails: people(r.ProjectMembers),
    description: str(r.Details).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(),
    color: str(r.ColorHex)
  };
}

export function toTmTask(r: Row): ITmTask {
  return {
    id: num(r.Id),
    title: str(r.Title),
    number: str(r.TaskNumber),
    projectId: num(r.ProjectId),
    statusKind: str(r.StatusKind) || 'NotStarted',
    assigneeEmail: str(r.AssigneeEmail),
    due: tmDay(r.DueDate),
    estimateHours: num(r.EstimateHours),
    spentMinutes: num(r.SpentMinutes)
  };
}

export function toTmTimeEntry(r: Row): ITmTimeEntry {
  return {
    id: num(r.Id),
    taskId: num(r.TaskId),
    projectId: num(r.ProjectId),
    minutes: num(r.DurationMinutes),
    note: str(r.WorkNote),
    billable: r.Billable === true || r.Billable === 1 || r.Billable === '1',
    workDate: str(r.WorkDate).substring(0, 10),
    authorEmail: person(r.Author)
  };
}

export interface ISyncPlan<T> {
  create: Array<Partial<T>>;
  update: Array<{ id: number; patch: Partial<T> }>;
}

const same = (a: unknown, b: unknown): boolean => JSON.stringify(a) === JSON.stringify(b);

/** Only the fields that differ (so an unchanged project or task is not written). */
function diff<T>(current: T, wanted: Partial<T>): Partial<T> {
  const out: Partial<T> = {};
  (Object.keys(wanted) as Array<keyof T>).forEach((k) => {
    if (!same(current[k], wanted[k])) out[k] = wanted[k];
  });
  return out;
}

/**
 * Projects: link by TmProjectId, else adopt a timesheet project with the same code; create the rest. Task Manager owns
 * name, code, status, dates, budget and people; rates, billable and custom columns stay as set in Timesheet.
 */
export function planProjects(tm: ITmProject[], ts: IProject[], statuses: string[], palette: string[]): ISyncPlan<IProject> {
  const plan: ISyncPlan<IProject> = { create: [], update: [] };
  const taken: Record<number, boolean> = {};
  tm.forEach((t, i) => {
    let mine = ts.filter((p) => p.tmProjectId === t.id)[0];
    if (!mine && t.code) mine = ts.filter((p) => !p.tmProjectId && !taken[p.id] && lc(p.code) === lc(t.code))[0];
    if (!mine && statuses.length && statuses.indexOf(t.status) < 0) return;
    const managers = t.ownerEmail ? [t.ownerEmail] : [];
    const members = t.memberEmails.length ? t.memberEmails.concat(managers).filter((e, j, a) => a.map(lc).indexOf(lc(e)) === j) : [];
    const wanted: Partial<IProject> = {
      title: t.title,
      code: t.code,
      status: projectStatusFromTm(t.status),
      startDate: t.start,
      endDate: t.end,
      plannedHours: t.budgetHours,
      tmProjectId: t.id
    };
    if (managers.length) wanted.managerEmails = managers;
    wanted.memberEmails = members;
    if (mine) {
      taken[mine.id] = true;
      const patch = diff(mine, wanted);
      if (Object.keys(patch).length) plan.update.push({ id: mine.id, patch });
    } else {
      plan.create.push({
        ...wanted,
        client: '',
        description: t.description,
        programId: 0,
        managerEmails: managers,
        billable: true,
        hourlyCost: 0,
        billRate: 0,
        memberRates: {},
        color: t.color || palette[i % palette.length],
        custom: {},
        endNotices: []
      });
    }
  });
  return plan;
}

/** Tasks of linked projects: link by TmTaskId; Task Manager owns name, status, due date, estimate and assignee. */
export function planTasks(tm: ITmTask[], ts: ITask[], projects: IProject[]): ISyncPlan<ITask> {
  const plan: ISyncPlan<ITask> = { create: [], update: [] };
  const byTm: Record<number, IProject> = {};
  projects.forEach((p) => p.tmProjectId && (byTm[p.tmProjectId] = p));
  tm.forEach((t) => {
    const project = byTm[t.projectId];
    if (!project) return;
    const wanted: Partial<ITask> = {
      title: t.number && t.title.indexOf(t.number) < 0 ? `${t.title} (${t.number})` : t.title,
      projectId: project.id,
      status: taskStatusFromTm(t.statusKind),
      dueDate: t.due,
      plannedHours: t.estimateHours,
      assigneeEmails: t.assigneeEmail ? [t.assigneeEmail] : [],
      tmTaskId: t.id
    };
    const mine = ts.filter((x) => x.tmTaskId === t.id)[0];
    if (mine) {
      const patch = diff(mine, wanted);
      if (Object.keys(patch).length) plan.update.push({ id: mine.id, patch });
    } else if (wanted.status !== 'Done') {
      plan.create.push({ ...wanted, description: '', billable: project.billable, selfAssign: true, plannerTaskId: '', plannerPlanId: '' });
    }
  });
  return plan;
}

/**
 * Time a person logged in Task Manager, as lines for their week: one line per timesheet task (or project), hours per
 * day, notes joined. Entries already imported into the week are skipped; entries on tasks that are not linked are
 * counted in `skipped`.
 */
export function importLines(entries: ITmTimeEntry[], weekStart: DateKey, email: string, tasks: ITask[], projects: IProject[], alreadyImported: number[]): { lines: ITimeEntry[]; skipped: number } {
  const days = weekDays(weekStart);
  const lines: Record<string, ITimeEntry> = {};
  let skipped = 0;
  entries
    .filter((e) => days.indexOf(e.workDate) >= 0 && alreadyImported.indexOf(e.id) < 0 && e.minutes > 0)
    .forEach((e) => {
      const task = tasks.filter((t) => t.tmTaskId === e.taskId)[0];
      const project = task ? projects.filter((p) => p.id === task.projectId)[0] : projects.filter((p) => p.tmProjectId === e.projectId)[0];
      if (!project) {
        skipped++;
        return;
      }
      const key = `${project.id}|${task ? task.id : 0}|${e.billable ? 1 : 0}`;
      const line =
        lines[key] ||
        (lines[key] = {
          id: 0,
          sheetId: 0,
          userEmail: email,
          weekStart,
          projectId: project.id,
          taskId: task ? task.id : 0,
          activityId: 0,
          billable: e.billable && project.billable,
          hours: {},
          notes: {},
          total: 0,
          custom: {},
          costRate: 0,
          billRate: 0,
          tmSync: { imported: [] }
        });
      line.hours[e.workDate] = Math.round(((line.hours[e.workDate] || 0) + e.minutes / 60) * 100) / 100;
      if (e.note) line.notes[e.workDate] = line.notes[e.workDate] ? `${line.notes[e.workDate]}; ${e.note}` : e.note;
      (line.tmSync.imported as number[]).push(e.id);
    });
  const out = Object.keys(lines).map((k) => lines[k]);
  out.forEach((l) => (l.total = Math.round(Object.keys(l.hours).reduce((s, d) => s + l.hours[d], 0) * 100) / 100));
  return { lines: out, skipped };
}

/** What a line writes to Task Manager: its linked task and one entry per day with hours. Imported lines write nothing. */
export function pushItems(line: ITimeEntry, tasks: ITask[]): { tmTaskId: number; items: Array<{ workDate: DateKey; minutes: number; note: string }> } {
  if (line.tmSync && line.tmSync.imported && line.tmSync.imported.length) return { tmTaskId: 0, items: [] };
  const task = tasks.filter((t) => t.id === line.taskId)[0];
  if (!task || !task.tmTaskId) return { tmTaskId: 0, items: [] };
  const items = Object.keys(line.hours)
    .filter((d) => (line.hours[d] || 0) > 0)
    .sort()
    .map((d) => ({ workDate: d, minutes: Math.round(line.hours[d] * 60), note: line.notes[d] || '' }));
  return { tmTaskId: task.tmTaskId, items };
}

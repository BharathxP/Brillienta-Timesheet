// Report engine: flattens timesheet lines into one fact per person / line / day, then groups and sums them.
// Every report (project, team, billable, custom reports, dashboards, exports) uses this. Pure (tests/reports.test.cjs).

import { DateKey, Dimension, IActivity, IProgram, IProject, IReportFilter, ITask, ITimeEntry, ITimesheet, ITsUser, Measure } from '../models/types';
import { isoWeek, monthKey, formatMonth, formatDay, formatWeek, weekStartOf } from './dates';

export interface IFact {
  day: DateKey;
  week: DateKey;
  month: string;
  hours: number;
  billable: boolean;
  cost: number;
  revenue: number;
  projectId: number;
  taskId: number;
  activityId: number;
  userEmail: string;
  status: string;
  custom: Record<string, string | number | boolean>;
  note: string;
}

export interface ILookups {
  projects: IProject[];
  tasks: ITask[];
  activities: IActivity[];
  users: ITsUser[];
  programs: IProgram[];
  /** Rates for lines of weeks not submitted yet (submitted lines carry their own). */
  rateFor?: (entry: ITimeEntry) => { cost: number; bill: number };
}

const lc = (s: string): string => (s || '').toLowerCase();

export function toFacts(entries: ITimeEntry[], sheets: ITimesheet[], lookups: ILookups, filter: Partial<IReportFilter>, weekStartDay: number): IFact[] {
  const status: Record<number, string> = {};
  sheets.forEach((s) => (status[s.id] = s.status));
  const facts: IFact[] = [];
  const users = (filter.userEmails || []).map(lc);
  entries.forEach((e) => {
    if (filter.projectIds && filter.projectIds.length && filter.projectIds.indexOf(e.projectId) < 0) return;
    if (users.length && users.indexOf(lc(e.userEmail)) < 0) return;
    if (filter.billable === 'billable' && !e.billable) return;
    if (filter.billable === 'nonBillable' && e.billable) return;
    const st = status[e.sheetId] || 'Draft';
    if (filter.statuses && filter.statuses.length && filter.statuses.indexOf(st as ITimesheet['status']) < 0) return;
    let rates = { cost: e.costRate, bill: e.billRate };
    if (!(rates.cost > 0 || rates.bill > 0) && lookups.rateFor) rates = lookups.rateFor(e);
    Object.keys(e.hours).forEach((day) => {
      const h = Number(e.hours[day]) || 0;
      if (!h) return;
      if (filter.from && day < filter.from) return;
      if (filter.to && day > filter.to) return;
      facts.push({
        day,
        week: weekStartOf(day, weekStartDay),
        month: monthKey(day),
        hours: h,
        billable: e.billable,
        cost: h * (rates.cost || 0),
        revenue: e.billable ? h * (rates.bill || 0) : 0,
        projectId: e.projectId,
        taskId: e.taskId,
        activityId: e.activityId,
        userEmail: e.userEmail,
        status: st,
        custom: e.custom || {},
        note: (e.notes || {})[day] || ''
      });
    });
  });
  return facts;
}

export const DIMENSIONS: Array<{ key: Dimension; text: string }> = [
  { key: 'project', text: 'Project' },
  { key: 'client', text: 'Client' },
  { key: 'program', text: 'Programme' },
  { key: 'task', text: 'Task' },
  { key: 'activity', text: 'Activity' },
  { key: 'user', text: 'Person' },
  { key: 'department', text: 'Department' },
  { key: 'billable', text: 'Billable' },
  { key: 'status', text: 'Timesheet status' },
  { key: 'day', text: 'Day' },
  { key: 'week', text: 'Week' },
  { key: 'month', text: 'Month' }
];

export const MEASURES: Array<{ key: Measure; text: string; money?: boolean }> = [
  { key: 'hours', text: 'Hours' },
  { key: 'billableHours', text: 'Billable hours' },
  { key: 'nonBillableHours', text: 'Non-billable hours' },
  { key: 'cost', text: 'Cost', money: true },
  { key: 'revenue', text: 'Billable amount', money: true },
  { key: 'margin', text: 'Margin', money: true },
  { key: 'entries', text: 'Entries' }
];

/** Group key and display label of a fact for a dimension. Custom columns are "custom:<key>". */
export function dimensionValue(f: IFact, dim: Dimension, lk: ILookups): { key: string; label: string; sort: string } {
  const project = (): IProject | undefined => lk.projects.filter((p) => p.id === f.projectId)[0];
  switch (dim) {
    case 'project': {
      const p = project();
      return { key: String(f.projectId), label: p ? (p.code ? `${p.code} · ${p.title}` : p.title) : '(no project)', sort: p ? p.title : '~' };
    }
    case 'client': {
      const p = project();
      const c = (p && p.client) || '(no client)';
      return { key: c, label: c, sort: c };
    }
    case 'program': {
      const p = project();
      const g = p && lk.programs.filter((x) => x.id === p.programId)[0];
      return { key: g ? String(g.id) : '0', label: g ? g.title : '(no programme)', sort: g ? g.title : '~' };
    }
    case 'task': {
      const t = lk.tasks.filter((x) => x.id === f.taskId)[0];
      return { key: String(f.taskId), label: t ? t.title : '(no task)', sort: t ? t.title : '~' };
    }
    case 'activity': {
      const a = lk.activities.filter((x) => x.id === f.activityId)[0];
      return { key: String(f.activityId), label: a ? a.title : '(no activity)', sort: a ? a.title : '~' };
    }
    case 'user': {
      const u = lk.users.filter((x) => lc(x.email) === lc(f.userEmail))[0];
      return { key: lc(f.userEmail), label: u ? u.name : f.userEmail, sort: u ? u.name : f.userEmail };
    }
    case 'department': {
      const u = lk.users.filter((x) => lc(x.email) === lc(f.userEmail))[0];
      const d = (u && u.department) || '(no department)';
      return { key: d, label: d, sort: d };
    }
    case 'billable':
      return { key: f.billable ? 'b' : 'n', label: f.billable ? 'Billable' : 'Non-billable', sort: f.billable ? 'a' : 'b' };
    case 'status':
      return { key: f.status, label: f.status, sort: f.status };
    case 'day':
      return { key: f.day, label: formatDay(f.day, true), sort: f.day };
    case 'week':
      return { key: f.week, label: `W${isoWeek(f.week)} · ${formatWeek(f.week)}`, sort: f.week };
    case 'month':
      return { key: f.month, label: formatMonth(f.month), sort: f.month };
    default: {
      if (String(dim).indexOf('custom:') === 0) {
        const v = f.custom[String(dim).substring(7)];
        const s = v === undefined || v === '' ? '(empty)' : typeof v === 'boolean' ? (v ? 'Yes' : 'No') : String(v);
        return { key: s, label: s, sort: s };
      }
      return { key: '', label: '', sort: '' };
    }
  }
}

export interface IGroupRow {
  keys: string[];
  labels: string[];
  values: Record<Measure, number>;
}

function measureOf(f: IFact, m: Measure): number {
  switch (m) {
    case 'hours':
      return f.hours;
    case 'billableHours':
      return f.billable ? f.hours : 0;
    case 'nonBillableHours':
      return f.billable ? 0 : f.hours;
    case 'cost':
      return f.cost;
    case 'revenue':
      return f.revenue;
    case 'margin':
      return f.revenue - f.cost;
    case 'entries':
      return 1;
    default:
      return 0;
  }
}

const ALL_MEASURES: Measure[] = ['hours', 'billableHours', 'nonBillableHours', 'cost', 'revenue', 'margin', 'entries'];

/** Group facts by the dimensions (in order) and sum every measure. Rows are sorted by the dimensions' sort keys. */
export function groupFacts(facts: IFact[], dims: Dimension[], lk: ILookups): IGroupRow[] {
  const map: Record<string, IGroupRow & { sorts: string[] }> = {};
  facts.forEach((f) => {
    const vals = dims.map((d) => dimensionValue(f, d, lk));
    const id = vals.map((v) => v.key).join('\u0001');
    let row = map[id];
    if (!row) {
      const zero = {} as Record<Measure, number>;
      ALL_MEASURES.forEach((m) => (zero[m] = 0));
      row = map[id] = { keys: vals.map((v) => v.key), labels: vals.map((v) => v.label), sorts: vals.map((v) => v.sort), values: zero };
    }
    ALL_MEASURES.forEach((m) => (row.values[m] += measureOf(f, m)));
  });
  const rows = Object.keys(map).map((k) => map[k]);
  rows.sort((a, b) => {
    for (let i = 0; i < a.sorts.length; i++) {
      const c = a.sorts[i].localeCompare(b.sorts[i]);
      if (c) return c;
    }
    return 0;
  });
  return rows.map((r) => ({ keys: r.keys, labels: r.labels, values: roundValues(r.values) }));
}

function roundValues(v: Record<Measure, number>): Record<Measure, number> {
  const out = {} as Record<Measure, number>;
  (Object.keys(v) as Measure[]).forEach((k) => (out[k] = Math.round(v[k] * 100) / 100));
  return out;
}

export function totals(facts: IFact[]): Record<Measure, number> {
  const out = {} as Record<Measure, number>;
  ALL_MEASURES.forEach((m) => (out[m] = facts.reduce((s, f) => s + measureOf(f, m), 0)));
  return roundValues(out);
}

/** Hours already used per project (all time) — planned hours tracking. */
export function usedByProject(entries: ITimeEntry[]): Record<number, number> {
  const out: Record<number, number> = {};
  entries.forEach((e) => (out[e.projectId] = Math.round(((out[e.projectId] || 0) + (e.total || 0)) * 100) / 100));
  return out;
}

/** Utilisation: billable hours / available hours (working days × standard day) in the period. */
export function utilisation(billableHours: number, availableHours: number): number {
  return availableHours > 0 ? Math.round((billableHours / availableHours) * 1000) / 10 : 0;
}

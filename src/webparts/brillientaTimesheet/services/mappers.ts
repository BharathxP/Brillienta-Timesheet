// SharePoint list rows ⇄ domain objects. Pure functions (tests/mappers.test.cjs).

import { IActivity, IAuditEntry, INotification, IProgram, IProject, ISavedReport, ISheetComment, ITask, ITicket, ITimeEntry, ITimesheet, ITsUser, Role, ROLES } from '../models/types';

export type Row = Record<string, unknown>;

const str = (v: unknown): string => (v === undefined || v === null ? '' : String(v));
const num = (v: unknown): number => {
  const n = Number(v);
  return isFinite(n) ? n : 0;
};
const bool = (v: unknown, dflt: boolean = false): boolean => (v === undefined || v === null ? dflt : v === true || v === 1 || v === '1' || v === 'true' || v === 'Yes');

export function json<T>(v: unknown, fallback: T): T {
  if (typeof v !== 'string' || !v) return fallback;
  try {
    const parsed = JSON.parse(v);
    return parsed === null || parsed === undefined ? fallback : (parsed as T);
  } catch {
    return fallback;
  }
}

/** Emails kept as "a@x.com; b@y.com" (readable in the list and in Power BI). */
export function emails(v: unknown): string[] {
  return str(v)
    .split(/[;,\n]+/)
    .map((s) => s.trim())
    .filter((s) => !!s);
}
export const joinEmails = (list: string[]): string => list.filter((x) => !!x).join('; ');

/** A date-only column comes back as an ISO instant at the site's midnight; keep the calendar day. */
export function dateOnly(v: unknown): string {
  const s = str(v);
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  if (isNaN(d.getTime())) return '';
  // SharePoint returns the UTC instant of local midnight on the site; noon-shift to get the right day in any zone.
  const shifted = new Date(d.getTime() + 12 * 3600000);
  return shifted.toISOString().substring(0, 10);
}
/** Write a date-only value (null clears it). */
export const dateOut = (key: string): string | null => (key ? key + 'T12:00:00Z' : null);

const author = (r: Row): string => {
  const a = r.Author as { Title?: string } | undefined;
  return (a && a.Title) || '';
};

export function toUser(r: Row): ITsUser {
  const role = str(r.UserRole) as Role;
  return {
    id: num(r.Id),
    name: str(r.Title),
    email: str(r.Email),
    role: ROLES.indexOf(role) >= 0 ? role : 'member',
    managerEmail: str(r.ManagerEmail),
    department: str(r.Department),
    jobTitle: str(r.JobTitle),
    costRate: num(r.CostRate),
    billRate: num(r.BillRate),
    external: bool(r.IsExternal),
    active: bool(r.IsActive, true),
    ruleId: str(r.HourRuleId),
    workDays: str(r.WorkDays)
      .split(',')
      .map((s) => s.trim())
      .filter((s) => /^[0-6]$/.test(s))
      .map(Number),
    created: str(r.Created)
  };
}

export function fromUser(u: Partial<ITsUser>): Row {
  const out: Row = {};
  if (u.name !== undefined) out.Title = u.name;
  if (u.email !== undefined) out.Email = u.email;
  if (u.role !== undefined) out.UserRole = u.role;
  if (u.managerEmail !== undefined) out.ManagerEmail = u.managerEmail;
  if (u.department !== undefined) out.Department = u.department;
  if (u.jobTitle !== undefined) out.JobTitle = u.jobTitle;
  if (u.costRate !== undefined) out.CostRate = u.costRate;
  if (u.billRate !== undefined) out.BillRate = u.billRate;
  if (u.external !== undefined) out.IsExternal = u.external;
  if (u.active !== undefined) out.IsActive = u.active;
  if (u.ruleId !== undefined) out.HourRuleId = u.ruleId;
  if (u.workDays !== undefined) out.WorkDays = u.workDays.join(',');
  return out;
}

export function toProgram(r: Row): IProgram {
  return { id: num(r.Id), title: str(r.Title), description: str(r.Details), managerEmails: emails(r.ManagerEmails), color: str(r.ColorHex) };
}

export function fromProgram(p: Partial<IProgram>): Row {
  const out: Row = {};
  if (p.title !== undefined) out.Title = p.title;
  if (p.description !== undefined) out.Details = p.description;
  if (p.managerEmails !== undefined) out.ManagerEmails = joinEmails(p.managerEmails);
  if (p.color !== undefined) out.ColorHex = p.color;
  return out;
}

export function toProject(r: Row): IProject {
  const status = str(r.ProjectStatus);
  return {
    id: num(r.Id),
    title: str(r.Title),
    code: str(r.ProjectCode),
    client: str(r.Client),
    description: str(r.Details),
    programId: num(r.ProgramId),
    managerEmails: emails(r.ManagerEmails),
    memberEmails: emails(r.MemberEmails),
    billable: bool(r.IsBillable, true),
    status: status === 'Disabled' || status === 'Completed' ? status : 'Active',
    startDate: dateOnly(r.StartDate),
    endDate: dateOnly(r.EndDate),
    plannedHours: num(r.PlannedHours),
    hourlyCost: num(r.HourlyCost),
    billRate: num(r.BillRate),
    memberRates: json<Record<string, number>>(r.MemberRates, {}),
    color: str(r.ColorHex),
    custom: json<Record<string, string | number | boolean>>(r.CustomValues, {}),
    endNotices: str(r.EndNotices)
      .split(',')
      .filter((s) => !!s),
    tmProjectId: num(r.TmProjectId),
    modified: str(r.Modified)
  };
}

export function fromProject(p: Partial<IProject>): Row {
  const out: Row = {};
  if (p.title !== undefined) out.Title = p.title;
  if (p.code !== undefined) out.ProjectCode = p.code;
  if (p.client !== undefined) out.Client = p.client;
  if (p.description !== undefined) out.Details = p.description;
  if (p.programId !== undefined) out.ProgramId = p.programId;
  if (p.managerEmails !== undefined) out.ManagerEmails = joinEmails(p.managerEmails);
  if (p.memberEmails !== undefined) out.MemberEmails = joinEmails(p.memberEmails);
  if (p.billable !== undefined) out.IsBillable = p.billable;
  if (p.status !== undefined) out.ProjectStatus = p.status;
  if (p.startDate !== undefined) out.StartDate = dateOut(p.startDate);
  if (p.endDate !== undefined) out.EndDate = dateOut(p.endDate);
  if (p.plannedHours !== undefined) out.PlannedHours = p.plannedHours;
  if (p.hourlyCost !== undefined) out.HourlyCost = p.hourlyCost;
  if (p.billRate !== undefined) out.BillRate = p.billRate;
  if (p.memberRates !== undefined) out.MemberRates = JSON.stringify(p.memberRates);
  if (p.color !== undefined) out.ColorHex = p.color;
  if (p.custom !== undefined) out.CustomValues = JSON.stringify(p.custom);
  if (p.endNotices !== undefined) out.EndNotices = p.endNotices.join(',');
  if (p.tmProjectId !== undefined) out.TmProjectId = p.tmProjectId;
  return out;
}

export function toTask(r: Row): ITask {
  const s = str(r.TaskStatus);
  return {
    id: num(r.Id),
    title: str(r.Title),
    projectId: num(r.ProjectId),
    description: str(r.Details),
    assigneeEmails: emails(r.AssigneeEmails),
    billable: bool(r.IsBillable, true),
    plannedHours: num(r.PlannedHours),
    status: s === 'In progress' || s === 'Done' ? s : 'Open',
    dueDate: dateOnly(r.DueDate),
    selfAssign: bool(r.SelfAssign, true),
    plannerTaskId: str(r.PlannerTaskId),
    plannerPlanId: str(r.PlannerPlanId),
    tmTaskId: num(r.TmTaskId)
  };
}

export function fromTask(t: Partial<ITask>): Row {
  const out: Row = {};
  if (t.title !== undefined) out.Title = t.title;
  if (t.projectId !== undefined) out.ProjectId = t.projectId;
  if (t.description !== undefined) out.Details = t.description;
  if (t.assigneeEmails !== undefined) out.AssigneeEmails = joinEmails(t.assigneeEmails);
  if (t.billable !== undefined) out.IsBillable = t.billable;
  if (t.plannedHours !== undefined) out.PlannedHours = t.plannedHours;
  if (t.status !== undefined) out.TaskStatus = t.status;
  if (t.dueDate !== undefined) out.DueDate = dateOut(t.dueDate);
  if (t.selfAssign !== undefined) out.SelfAssign = t.selfAssign;
  if (t.plannerTaskId !== undefined) out.PlannerTaskId = t.plannerTaskId;
  if (t.plannerPlanId !== undefined) out.PlannerPlanId = t.plannerPlanId;
  if (t.tmTaskId !== undefined) out.TmTaskId = t.tmTaskId;
  return out;
}

export function toActivity(r: Row): IActivity {
  return { id: num(r.Id), title: str(r.Title), billable: bool(r.IsBillable, true), active: bool(r.IsActive, true), color: str(r.ColorHex), isTimeOff: bool(r.IsTimeOff) };
}

export function fromActivity(a: Partial<IActivity>): Row {
  const out: Row = {};
  if (a.title !== undefined) out.Title = a.title;
  if (a.billable !== undefined) out.IsBillable = a.billable;
  if (a.active !== undefined) out.IsActive = a.active;
  if (a.color !== undefined) out.ColorHex = a.color;
  if (a.isTimeOff !== undefined) out.IsTimeOff = a.isTimeOff;
  return out;
}

export function sheetKey(email: string, weekStart: string): string {
  return `${email.toLowerCase()}|${weekStart}`;
}

export function toSheet(r: Row): ITimesheet {
  const s = str(r.SheetStatus);
  return {
    id: num(r.Id),
    userEmail: str(r.UserEmail),
    userName: str(r.UserName),
    weekStart: str(r.WeekStart),
    status: s === 'Submitted' || s === 'Approved' || s === 'Rejected' ? s : 'Draft',
    totalHours: num(r.TotalHours),
    billableHours: num(r.BillableHours),
    submittedAt: str(r.SubmittedAt),
    decidedBy: str(r.DecidedBy),
    decidedAt: str(r.DecidedAt),
    autoApproved: bool(r.AutoApproved),
    comments: json<ISheetComment[]>(r.Comments, []),
    modified: str(r.Modified)
  };
}

export function fromSheet(s: Partial<ITimesheet>): Row {
  const out: Row = {};
  if (s.userEmail !== undefined && s.weekStart !== undefined) out.Title = sheetKey(s.userEmail, s.weekStart);
  if (s.userEmail !== undefined) out.UserEmail = s.userEmail;
  if (s.userName !== undefined) out.UserName = s.userName;
  if (s.weekStart !== undefined) out.WeekStart = s.weekStart;
  if (s.status !== undefined) out.SheetStatus = s.status;
  if (s.totalHours !== undefined) out.TotalHours = s.totalHours;
  if (s.billableHours !== undefined) out.BillableHours = s.billableHours;
  if (s.submittedAt !== undefined) out.SubmittedAt = s.submittedAt || null;
  if (s.decidedBy !== undefined) out.DecidedBy = s.decidedBy;
  if (s.decidedAt !== undefined) out.DecidedAt = s.decidedAt || null;
  if (s.autoApproved !== undefined) out.AutoApproved = s.autoApproved;
  if (s.comments !== undefined) out.Comments = JSON.stringify(s.comments.slice(-50));
  return out;
}

export function toEntry(r: Row): ITimeEntry {
  const hours = json<Record<string, number>>(r.DayHours, {});
  Object.keys(hours).forEach((k) => (hours[k] = num(hours[k])));
  return {
    id: num(r.Id),
    sheetId: num(r.SheetId),
    userEmail: str(r.UserEmail),
    weekStart: str(r.WeekStart),
    projectId: num(r.ProjectId),
    taskId: num(r.TaskId),
    activityId: num(r.ActivityId),
    billable: bool(r.IsBillable, true),
    hours,
    notes: json<Record<string, string>>(r.DayNotes, {}),
    total: num(r.TotalHours),
    custom: json<Record<string, string | number | boolean>>(r.CustomValues, {}),
    costRate: num(r.CostRate),
    billRate: num(r.BillRate),
    tmSync: json(r.TmSync, {})
  };
}

export function entryTotal(e: Pick<ITimeEntry, 'hours'>): number {
  return Math.round(Object.keys(e.hours).reduce((s, k) => s + (Number(e.hours[k]) || 0), 0) * 100) / 100;
}

export function fromEntry(e: ITimeEntry): Row {
  const hours: Record<string, number> = {};
  Object.keys(e.hours).forEach((k) => {
    if (e.hours[k]) hours[k] = e.hours[k];
  });
  const notes: Record<string, string> = {};
  Object.keys(e.notes).forEach((k) => {
    if (e.notes[k]) notes[k] = e.notes[k];
  });
  return {
    Title: `${e.userEmail.toLowerCase()}|${e.weekStart}|${e.projectId}|${e.taskId}|${e.activityId}`.substring(0, 255),
    SheetId: e.sheetId,
    UserEmail: e.userEmail,
    WeekStart: e.weekStart,
    ProjectId: e.projectId,
    TaskId: e.taskId,
    ActivityId: e.activityId,
    IsBillable: e.billable,
    DayHours: JSON.stringify(hours),
    DayNotes: JSON.stringify(notes),
    TotalHours: entryTotal({ hours }),
    CustomValues: JSON.stringify(e.custom || {}),
    CostRate: e.costRate || 0,
    BillRate: e.billRate || 0,
    TmSync: JSON.stringify(e.tmSync || {})
  };
}

export function toTicket(r: Row): ITicket {
  const p = str(r.Priority);
  const s = str(r.TicketStatus);
  return {
    id: num(r.Id),
    title: str(r.Title),
    details: str(r.Details),
    category: str(r.Category),
    priority: p === 'Low' || p === 'High' || p === 'Urgent' ? p : 'Normal',
    status: s === 'In progress' || s === 'Answered' || s === 'Closed' ? s : 'Open',
    raisedBy: author(r),
    raisedByEmail: str(r.RaisedByEmail),
    created: str(r.Created),
    reply: str(r.Reply)
  };
}

export function toNotification(r: Row): INotification {
  return { id: num(r.Id), title: str(r.Title), body: str(r.Body), link: str(r.LinkPath), recipient: str(r.Recipient), read: bool(r.IsRead), created: str(r.Created), kind: str(r.Kind) };
}

export function toAudit(r: Row): IAuditEntry {
  return { id: num(r.Id), action: str(r.Title), details: str(r.Details), by: str(r.ByName) || author(r), created: str(r.Created) };
}

export function toReport(r: Row): ISavedReport {
  return {
    id: num(r.Id),
    title: str(r.Title),
    owner: str(r.OwnerEmail),
    shared: bool(r.IsShared),
    kind: str(r.ReportKind) === 'dashboard' ? 'dashboard' : 'report',
    definition: json(r.Definition, { groupBy: ['project'], measures: ['hours'], chart: 'table', period: 'thisMonth', filter: { from: '', to: '', projectIds: [], userEmails: [], billable: 'all', statuses: [] } })
  };
}

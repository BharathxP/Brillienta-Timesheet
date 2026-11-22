// Domain model of Brillienta Timesheet. Every list row is mapped to one of these types in services/mappers.ts.

/**
 * The four standard roles plus Program Manager (Pro and Enterprise).
 * - admin: everything, settings included.
 * - programManager: manages the projects of their programmes, approves and reports on them.
 * - projectManager: creates and runs their own projects (tasks, members, budgets), approves time on them.
 * - teamManager: approves the timesheets of the people who report to them; "My team" report.
 * - member: fills in their own timesheet, self-assigns tasks.
 */
export type Role = 'admin' | 'programManager' | 'projectManager' | 'teamManager' | 'member';

export const ROLES: Role[] = ['admin', 'programManager', 'projectManager', 'teamManager', 'member'];

export type SheetStatus = 'Draft' | 'Submitted' | 'Approved' | 'Rejected';

export type ProjectStatus = 'Active' | 'Disabled' | 'Completed';

export type TaskStatus = 'Open' | 'In progress' | 'Done';

/** yyyy-mm-dd (local calendar day). */
export type DateKey = string;

export interface ICurrentUser {
  id: number;
  title: string;
  email: string;
  loginName: string;
  isSiteAdmin: boolean;
  canManageLists: boolean;
  canManageWeb: boolean;
  isExternal: boolean;
}

/** A person who tracks time (TS_Users). */
export interface ITsUser {
  id: number;
  name: string;
  email: string;
  role: Role;
  managerEmail: string;
  department: string;
  jobTitle: string;
  /** What an hour of this person costs the company (user-specific hourly cost). 0 = not set. */
  costRate: number;
  /** What an hour of this person is billed at (user-specific rate). 0 = not set. */
  billRate: number;
  external: boolean;
  active: boolean;
  /** Hour rule that applies (Enterprise, rule-based restrictions); '' = matched by the rules' conditions. */
  ruleId: string;
  /** Days this person tracks time (0 = Sunday … 6 = Saturday); empty = the organisation's working days. */
  workDays: number[];
  created?: string;
}

/** A group of projects run by one or more programme managers. */
export interface IProgram {
  id: number;
  title: string;
  description: string;
  managerEmails: string[];
  color: string;
}

export interface IProject {
  id: number;
  title: string;
  code: string;
  client: string;
  description: string;
  programId: number;
  managerEmails: string[];
  /** Who can log time; empty = everyone. */
  memberEmails: string[];
  billable: boolean;
  status: ProjectStatus;
  startDate: DateKey;
  endDate: DateKey;
  /** Planned / budgeted hours (0 = not planned). */
  plannedHours: number;
  /** Flat hourly cost of the project (used when a person has no cost rate of their own, or always — see costMode). */
  hourlyCost: number;
  /** Hourly bill rate of the project. */
  billRate: number;
  /** Person-specific bill rates on this project: email → rate. */
  memberRates: Record<string, number>;
  color: string;
  /** Values of the custom project columns. */
  custom: Record<string, string | number | boolean>;
  /** End-date notifications already sent ("7", "0", "overdue"). */
  endNotices: string[];
  modified?: string;
}

export interface ITask {
  id: number;
  title: string;
  projectId: number;
  description: string;
  assigneeEmails: string[];
  billable: boolean;
  plannedHours: number;
  status: TaskStatus;
  dueDate: DateKey;
  /** Anyone allowed on the project may assign it to themself. */
  selfAssign: boolean;
  plannerTaskId: string;
  plannerPlanId: string;
}

/** A kind of work (development, meetings, testing, travel…), for activity tracking. */
export interface IActivity {
  id: number;
  title: string;
  billable: boolean;
  active: boolean;
  color: string;
  /** Reserved activity for approved time off (Time off Manager 365). */
  isTimeOff: boolean;
}

export interface ISheetComment {
  at: string;
  by: string;
  byEmail: string;
  text: string;
  action?: 'submitted' | 'approved' | 'rejected' | 'reopened' | 'auto-approved' | 'comment';
}

/** One person's week (TS_Timesheets). The hours themselves are on the entries. */
export interface ITimesheet {
  id: number;
  userEmail: string;
  userName: string;
  weekStart: DateKey;
  status: SheetStatus;
  totalHours: number;
  billableHours: number;
  submittedAt: string;
  decidedBy: string;
  decidedAt: string;
  autoApproved: boolean;
  comments: ISheetComment[];
  modified?: string;
}

/** One line of a week: a project / task / activity with hours per day (TS_TimeEntries). */
export interface ITimeEntry {
  id: number;
  sheetId: number;
  userEmail: string;
  weekStart: DateKey;
  projectId: number;
  taskId: number;
  activityId: number;
  billable: boolean;
  /** Hours per day: yyyy-mm-dd → hours. */
  hours: Record<DateKey, number>;
  /** Note per day. */
  notes: Record<DateKey, string>;
  total: number;
  /** Values of the custom entry columns. */
  custom: Record<string, string | number | boolean>;
  /** Rates captured when the week was submitted (so later rate changes do not rewrite history). */
  costRate: number;
  billRate: number;
}

export type TicketStatus = 'Open' | 'In progress' | 'Answered' | 'Closed';

export interface ITicket {
  id: number;
  title: string;
  details: string;
  category: string;
  priority: 'Low' | 'Normal' | 'High' | 'Urgent';
  status: TicketStatus;
  raisedBy: string;
  raisedByEmail: string;
  created: string;
  reply: string;
}

export interface INotification {
  id: number;
  title: string;
  body: string;
  link: string;
  recipient: string;
  read: boolean;
  created: string;
  kind: string;
}

export interface IAuditEntry {
  id: number;
  action: string;
  details: string;
  by: string;
  created: string;
}

/** Approved time off read from Time off Manager 365 (or any leave list with the mapped columns). */
export interface ITimeOff {
  email: string;
  start: DateKey;
  end: DateKey;
  type: string;
  /** Hours per day off (full day = the day's standard hours; half days come from the leave list when mapped). */
  hoursPerDay?: number;
}

// --------------------------------------------------------------------------------------------- configuration

export type CustomColumnType = 'text' | 'number' | 'choice' | 'boolean' | 'date';

export interface ICustomColumn {
  key: string;
  label: string;
  type: CustomColumnType;
  /** 'entry': a column of each timesheet line; 'project': a field of each project. */
  level: 'entry' | 'project';
  choices: string[];
  required: boolean;
  /** Shown in reports as a group-by / filter. */
  inReports: boolean;
}

/** Daily / weekly hour limits. 0 = no limit. */
export interface IHourLimits {
  dailyMin: number;
  dailyMax: number;
  weeklyMin: number;
  weeklyMax: number;
}

/** Rule-based hour restriction (Enterprise): the first rule whose conditions match a person applies. */
export interface IHourRule extends IHourLimits {
  id: string;
  name: string;
  /** Conditions (empty = any). */
  roles: Role[];
  departments: string[];
  emails: string[];
  externalOnly: boolean;
  /** May log time on days that are not working days. */
  allowNonWorkingDays: boolean;
  /** Block submit ('block') or only warn ('warn'). */
  enforcement: 'block' | 'warn';
}

export interface IAutoApproval {
  enabled: boolean;
  /** Roles whose timesheets are approved as soon as they are submitted. */
  roles: Role[];
  /** Only when the week passes every hour rule without warnings. */
  requireWithinRules: boolean;
  /** Only when the week has at most this many hours (0 = any). */
  maxHours: number;
}

export type TemplateKey =
  | 'sheetSubmitted'
  | 'sheetApproved'
  | 'sheetRejected'
  | 'reminder'
  | 'projectEnding'
  | 'projectEnded'
  | 'taskAssigned'
  | 'plannedHoursExceeded'
  | 'ticketRaised';

export interface IEmailTemplate {
  subject: string;
  body: string;
  /** Also post the message to the person in Teams (Teams connectivity). */
  teams: boolean;
  enabled: boolean;
}

export interface IReminderConfig {
  enabled: boolean;
  /** Day of the week the reminder for the past week goes out (0 = Sunday). */
  day: number;
  /** Hour of the day (local, 0–23) from which it may go out. */
  hour: number;
  /** Also remind about the current week on its last working day. */
  currentWeek: boolean;
  /** Copy the person's manager. */
  ccManager: boolean;
  /** Weeks back that are still chased. */
  lookbackWeeks: number;
}

export interface ITimeOffConfig {
  enabled: boolean;
  /** Site of Time off Manager 365 ('' = this site). */
  siteUrl: string;
  listTitle: string;
  emailField: string;
  startField: string;
  endField: string;
  statusField: string;
  approvedValue: string;
  typeField: string;
  /** Put approved leave on the timesheet automatically (on the time-off activity). */
  autoFill: boolean;
}

export interface IIntegrationsConfig {
  /** How emails are sent: Outlook (Microsoft Graph, as the signed-in user) or the email queue + Power Automate. */
  mailMode: 'graph' | 'flow';
  /** Post notifications to people in Teams chat (Microsoft Graph). */
  teamsChat: boolean;
  plannerEnabled: boolean;
  /** Mark Planner tasks complete when their timesheet task is done. */
  plannerWriteBack: boolean;
  /** Power Automate "When an HTTP request is received" URL; the app posts events to it. */
  webhookUrl: string;
  webhookEvents: string[];
  copilotEnabled: boolean;
}

export type CostMode = 'userFirst' | 'projectFirst' | 'userOnly' | 'projectOnly';

export interface ITimesheetConfig {
  orgName: string;
  appName: string;
  brandColor: string;
  currency: string;
  /** Week starts on (0 = Sunday … 6 = Saturday). */
  weekStart: number;
  /** Days time is tracked on. */
  workDays: number[];
  /** Dates without tracking (public holidays…). */
  holidays: Array<{ date: DateKey; name: string }>;
  /** A standard working day (hours), used for leave and utilisation. */
  standardDay: number;
  /** Hour granularity (0.25 = quarter hours). */
  increment: number;
  limits: IHourLimits & { enforcement: 'block' | 'warn' };
  rulesEnabled: boolean;
  rules: IHourRule[];
  autoApproval: IAutoApproval;
  /** Who approves: the person's manager, the projects' managers, or either. */
  approvalBy: 'manager' | 'project' | 'either';
  /** People may change billable on a line (otherwise it follows project / task / activity). */
  allowBillableOverride: boolean;
  /** Require a note on every day with hours. */
  requireNotes: boolean;
  /** Activities are required on each line (activity tracking). */
  requireActivity: boolean;
  /** People may assign open self-assign tasks to themselves. */
  selfAssign: boolean;
  /** Members can create tasks on projects they belong to. */
  membersCreateTasks: boolean;
  costMode: CostMode;
  defaultCostRate: number;
  defaultBillRate: number;
  customColumns: ICustomColumn[];
  reminders: IReminderConfig;
  /** Days before a project's end date the project managers are told (Enterprise). */
  projectEndNotice: number[];
  /** Tell project managers when logged hours pass the planned hours. */
  plannedHoursAlerts: boolean;
  templates: Record<TemplateKey, IEmailTemplate>;
  timeOff: ITimeOffConfig;
  integrations: IIntegrationsConfig;
  /** Extra admins (besides site owners and people with the Administrator role). */
  admins: string[];
  /** New people start with this role. */
  defaultRole: Role;
  /** People may open the app without being added first (they are added as members on first visit). */
  autoEnrol: boolean;
  support: { email: string; accountManager: string; accountManagerEmail: string; phone: string };
}

export interface ISavedReport {
  id: number;
  title: string;
  owner: string;
  shared: boolean;
  /** 'report' = one table / chart; 'dashboard' = several widgets. */
  kind: 'report' | 'dashboard';
  definition: IReportDefinition | IDashboardDefinition;
}

export type Dimension = 'project' | 'client' | 'program' | 'task' | 'activity' | 'user' | 'department' | 'billable' | 'week' | 'month' | 'day' | 'status' | string;
export type Measure = 'hours' | 'billableHours' | 'nonBillableHours' | 'cost' | 'revenue' | 'margin' | 'entries';

export interface IReportFilter {
  from: DateKey;
  to: DateKey;
  projectIds: number[];
  userEmails: string[];
  billable: 'all' | 'billable' | 'nonBillable';
  statuses: SheetStatus[];
}

export interface IReportDefinition {
  groupBy: Dimension[];
  measures: Measure[];
  chart: 'table' | 'bar' | 'column' | 'donut' | 'line';
  filter: IReportFilter;
  /** Relative period, recomputed each time the report opens ('' = the saved dates). */
  period: '' | 'thisWeek' | 'lastWeek' | 'thisMonth' | 'lastMonth' | 'thisQuarter' | 'thisYear' | 'last30' | 'last90';
}

export interface IDashboardWidget {
  id: string;
  title: string;
  report: IReportDefinition;
  width: 1 | 2 | 3;
}

export interface IDashboardDefinition {
  widgets: IDashboardWidget[];
}

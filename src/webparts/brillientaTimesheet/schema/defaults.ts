import { IEmailTemplate, ITimesheetConfig, TemplateKey } from '../models/types';

export const DEFAULT_BRAND = '#17232e';

/** Standard email and Teams message templates. {{placeholders}} are filled in by services/templates.ts. */
export function defaultTemplates(): Record<TemplateKey, IEmailTemplate> {
  const t = (subject: string, body: string, teams: boolean = true): IEmailTemplate => ({ subject, body, teams, enabled: true });
  return {
    sheetSubmitted: t(
      'Timesheet to approve: {{person}}, week of {{week}}',
      '<p>Hello {{recipient}},</p><p><strong>{{person}}</strong> submitted the timesheet for the week of <strong>{{week}}</strong>: {{hours}} hours ({{billable}} billable).</p><p><a href="{{link}}">Review and approve</a></p>'
    ),
    sheetApproved: t('Your timesheet for the week of {{week}} is approved', '<p>Hello {{recipient}},</p><p>{{approver}} approved your timesheet for the week of <strong>{{week}}</strong> ({{hours}} hours).</p><p>{{comment}}</p><p><a href="{{link}}">Open the timesheet</a></p>'),
    sheetRejected: t(
      'Your timesheet for the week of {{week}} needs changes',
      '<p>Hello {{recipient}},</p><p>{{approver}} sent back your timesheet for the week of <strong>{{week}}</strong>.</p><blockquote>{{comment}}</blockquote><p><a href="{{link}}">Correct and resubmit</a></p>'
    ),
    reminder: t(
      'Reminder: your timesheet for the week of {{week}} is not submitted',
      '<p>Hello {{recipient}},</p><p>Your timesheet for the week of <strong>{{week}}</strong> has not been submitted yet ({{hours}} hours entered so far).</p><p><a href="{{link}}">Complete and submit it now</a></p>'
    ),
    projectEnding: t('Project {{project}} ends in {{days}} days', '<p>Hello {{recipient}},</p><p>The project <strong>{{project}}</strong> ends on <strong>{{endDate}}</strong>. {{used}} of {{planned}} planned hours are used.</p><p><a href="{{link}}">Open the project</a></p>'),
    projectEnded: t('Project {{project}} has reached its end date', '<p>Hello {{recipient}},</p><p>The project <strong>{{project}}</strong> reached its end date ({{endDate}}). Consider completing or extending it.</p><p><a href="{{link}}">Open the project</a></p>'),
    taskAssigned: t('New task: {{task}}', '<p>Hello {{recipient}},</p><p>You were assigned the task <strong>{{task}}</strong> on the project {{project}}.</p><p><a href="{{link}}">Open your tasks</a></p>'),
    plannedHoursExceeded: t('Project {{project}} passed its planned hours', '<p>Hello {{recipient}},</p><p><strong>{{project}}</strong> has used {{used}} hours of the {{planned}} planned.</p><p><a href="{{link}}">Open the project report</a></p>'),
    ticketRaised: t('Support ticket: {{subject}}', '<p>New support ticket from {{person}} ({{email}}) at {{org}}.</p><p><strong>{{subject}}</strong> — priority {{priority}}, category {{category}}</p><p>{{details}}</p>', false)
  };
}

export const TEMPLATE_LABELS: Record<TemplateKey, { label: string; when: string; placeholders: string[] }> = {
  sheetSubmitted: { label: 'Timesheet submitted', when: 'To the approvers, when someone submits a week.', placeholders: ['recipient', 'person', 'week', 'hours', 'billable', 'link'] },
  sheetApproved: { label: 'Timesheet approved', when: 'To the person, when their week is approved.', placeholders: ['recipient', 'approver', 'week', 'hours', 'comment', 'link'] },
  sheetRejected: { label: 'Timesheet rejected', when: 'To the person, when their week is sent back.', placeholders: ['recipient', 'approver', 'week', 'comment', 'link'] },
  reminder: { label: 'Pending timesheet reminder', when: 'To people whose week is not submitted (Admin › Notifications).', placeholders: ['recipient', 'week', 'hours', 'link'] },
  projectEnding: { label: 'Project end date approaching', when: 'To the project managers, the configured days before the end date.', placeholders: ['recipient', 'project', 'days', 'endDate', 'used', 'planned', 'link'] },
  projectEnded: { label: 'Project end date reached', when: 'To the project managers, on the end date.', placeholders: ['recipient', 'project', 'endDate', 'link'] },
  taskAssigned: { label: 'Task assigned', when: 'To the assignee, when someone else assigns them a task.', placeholders: ['recipient', 'task', 'project', 'link'] },
  plannedHoursExceeded: { label: 'Planned hours passed', when: 'To the project managers, when logged hours pass the planned hours.', placeholders: ['recipient', 'project', 'used', 'planned', 'link'] },
  ticketRaised: { label: 'Support ticket', when: 'To Brillienta support, when an admin raises a ticket.', placeholders: ['person', 'email', 'org', 'subject', 'priority', 'category', 'details'] }
};

export function defaultConfig(orgName: string = ''): ITimesheetConfig {
  return {
    orgName,
    appName: 'Brillienta Timesheet',
    brandColor: DEFAULT_BRAND,
    currency: 'INR',
    weekStart: 1,
    workDays: [1, 2, 3, 4, 5],
    holidays: [],
    standardDay: 8,
    increment: 0.25,
    limits: { dailyMin: 0, dailyMax: 24, weeklyMin: 0, weeklyMax: 0, enforcement: 'warn' },
    rulesEnabled: false,
    rules: [],
    autoApproval: { enabled: false, roles: ['admin'], requireWithinRules: true, maxHours: 0 },
    approvalBy: 'either',
    allowBillableOverride: true,
    requireNotes: false,
    requireActivity: false,
    selfAssign: true,
    membersCreateTasks: false,
    costMode: 'userFirst',
    defaultCostRate: 0,
    defaultBillRate: 0,
    customColumns: [],
    reminders: { enabled: true, day: 1, hour: 9, currentWeek: false, ccManager: false, lookbackWeeks: 2 },
    projectEndNotice: [14, 7, 1],
    plannedHoursAlerts: true,
    templates: defaultTemplates(),
    timeOff: {
      enabled: false,
      siteUrl: '',
      listTitle: 'TOM_LeaveRequests',
      emailField: 'EmployeeEmail',
      startField: 'StartDate',
      endField: 'EndDate',
      statusField: 'RequestStatus',
      approvedValue: 'Approved',
      typeField: 'LeaveType',
      autoFill: true
    },
    taskManager: {
      enabled: false,
      siteUrl: '',
      prefix: 'TM_',
      syncProjects: true,
      syncTasks: true,
      projectStatuses: ['Planning', 'Active', 'OnHold'],
      pushTime: 'approved',
      importTime: true,
      autoSync: true,
      lastSyncAt: ''
    },
    integrations: {
      mailMode: 'graph',
      teamsChat: false,
      plannerEnabled: false,
      plannerWriteBack: false,
      webhookUrl: '',
      webhookEvents: ['timesheet.submitted', 'timesheet.approved', 'timesheet.rejected'],
      copilotEnabled: false
    },
    admins: [],
    defaultRole: 'member',
    autoEnrol: true,
    support: { email: 'support@brillienta.com', accountManager: '', accountManagerEmail: '', phone: '' }
  };
}

function isObj(v: unknown): v is Record<string, unknown> {
  return !!v && typeof v === 'object' && !Array.isArray(v);
}

/** Deep merge of the saved configuration over the defaults, so settings added in later versions get their defaults. */
export function mergeConfig(saved: unknown): ITimesheetConfig {
  const base = defaultConfig() as unknown as Record<string, unknown>;
  const merge = (into: Record<string, unknown>, from: Record<string, unknown>): Record<string, unknown> => {
    Object.keys(from).forEach((k) => {
      const v = from[k];
      if (v === undefined || v === null) return;
      if (isObj(v) && isObj(into[k])) into[k] = merge({ ...(into[k] as Record<string, unknown>) }, v);
      else into[k] = v;
    });
    return into;
  };
  return (isObj(saved) ? merge(base, saved) : base) as unknown as ITimesheetConfig;
}

export const DEFAULT_ACTIVITIES: Array<{ title: string; billable: boolean; color: string; timeOff?: boolean }> = [
  { title: 'Development', billable: true, color: '#2a78d6' },
  { title: 'Design', billable: true, color: '#8a5cd1' },
  { title: 'Testing', billable: true, color: '#1baf7a' },
  { title: 'Meetings', billable: true, color: '#d18a2a' },
  { title: 'Project management', billable: true, color: '#17232e' },
  { title: 'Support', billable: true, color: '#c2410c' },
  { title: 'Training', billable: false, color: '#6b7280' },
  { title: 'Internal', billable: false, color: '#94a3b8' },
  { title: 'Time off', billable: false, color: '#9dc56a', timeOff: true }
];

export const PALETTE = ['#2a78d6', '#1baf7a', '#d18a2a', '#8a5cd1', '#c2410c', '#0e7490', '#b4237a', '#526d36', '#17232e', '#a16207'];

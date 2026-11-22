// Schema-as-code: every SharePoint list and column Brillienta Timesheet needs.
//
// The provisioner (services/Provisioner.ts) creates these on first install and, whenever the app version or
// SCHEMA_VERSION changes, re-validates them and adds anything missing. To add a column in a later release, add it here
// and bump SCHEMA_VERSION; existing sites pick it up the next time a site owner opens the app.
//
// People are stored by email (text), not as Person columns, so Microsoft 365 guests (external users) work the same as
// members and the data reads cleanly in Power BI, Power Automate and the REST API.

// 2: Brillienta Task Manager links (Projects.TmProjectId, Tasks.TmTaskId, TimeEntries.TmSync).
export const SCHEMA_VERSION = 2;

export type FieldType = 'Text' | 'Note' | 'Number' | 'Boolean' | 'DateTime' | 'DateOnly' | 'Choice';

export interface IFieldDef {
  name: string;
  title: string;
  type: FieldType;
  required?: boolean;
  choices?: string[];
  defaultValue?: string;
  indexed?: boolean;
  description?: string;
}

export type ListKey =
  | 'settings'
  | 'users'
  | 'programs'
  | 'projects'
  | 'tasks'
  | 'activities'
  | 'timesheets'
  | 'entries'
  | 'tickets'
  | 'notifications'
  | 'emailQueue'
  | 'audit'
  | 'reports';

/**
 * Who may do what on a list once security is applied (Admin › Schema & security):
 * - config: admins edit, everyone reads (settings, activities, programmes).
 * - managed: admins and managers edit, everyone reads (projects, tasks, people).
 * - timesheet: everyone adds and edits only their own rows; approvers and admins see and edit all.
 * - shared: everyone adds and edits (tasks — so people can self-assign and update progress —, notifications, tickets).
 * - adminOnly: admins only (audit log).
 */
export type SecurityProfile = 'config' | 'managed' | 'timesheet' | 'shared' | 'adminOnly';

export interface IListDef {
  key: ListKey;
  suffix: string;
  description: string;
  titleDisplayName?: string;
  versioning?: boolean;
  security: SecurityProfile;
  fields: IFieldDef[];
}

export const LISTS: IListDef[] = [
  {
    key: 'settings',
    suffix: 'Settings',
    description: 'Brillienta Timesheet configuration and install metadata (key/value).',
    titleDisplayName: 'Key',
    security: 'config',
    fields: [{ name: 'SettingValue', title: 'Value', type: 'Note' }]
  },
  {
    key: 'users',
    suffix: 'Users',
    description: 'People who track time: role, manager, hourly cost and bill rate.',
    titleDisplayName: 'Name',
    versioning: true,
    security: 'managed',
    fields: [
      { name: 'Email', title: 'Email', type: 'Text', indexed: true, required: true },
      { name: 'UserRole', title: 'Role', type: 'Choice', choices: ['admin', 'programManager', 'projectManager', 'teamManager', 'member'], defaultValue: 'member' },
      { name: 'ManagerEmail', title: 'Manager email', type: 'Text', indexed: true },
      { name: 'Department', title: 'Department', type: 'Text' },
      { name: 'JobTitle', title: 'Job title', type: 'Text' },
      { name: 'CostRate', title: 'Hourly cost', type: 'Number', defaultValue: '0' },
      { name: 'BillRate', title: 'Hourly bill rate', type: 'Number', defaultValue: '0' },
      { name: 'IsExternal', title: 'External user', type: 'Boolean', defaultValue: '0' },
      { name: 'IsActive', title: 'Active', type: 'Boolean', defaultValue: '1' },
      { name: 'HourRuleId', title: 'Hour rule', type: 'Text' },
      { name: 'WorkDays', title: 'Working days', type: 'Text', description: 'Comma-separated 0–6 (0 = Sunday); empty = the organisation default.' }
    ]
  },
  {
    key: 'programs',
    suffix: 'Programs',
    description: 'Programmes: groups of projects with programme managers.',
    titleDisplayName: 'Programme',
    security: 'config',
    fields: [
      { name: 'Details', title: 'Description', type: 'Note' },
      { name: 'ManagerEmails', title: 'Programme managers (emails)', type: 'Note' },
      { name: 'ColorHex', title: 'Colour', type: 'Text' }
    ]
  },
  {
    key: 'projects',
    suffix: 'Projects',
    description: 'Projects time is logged on: client, managers, members, dates, planned hours and rates.',
    titleDisplayName: 'Project',
    versioning: true,
    security: 'managed',
    fields: [
      { name: 'ProjectCode', title: 'Code', type: 'Text', indexed: true },
      { name: 'Client', title: 'Client', type: 'Text' },
      { name: 'Details', title: 'Description', type: 'Note' },
      { name: 'ProgramId', title: 'Programme ID', type: 'Number', defaultValue: '0' },
      { name: 'ManagerEmails', title: 'Project managers (emails)', type: 'Note' },
      { name: 'MemberEmails', title: 'Members (emails)', type: 'Note' },
      { name: 'IsBillable', title: 'Billable', type: 'Boolean', defaultValue: '1' },
      { name: 'ProjectStatus', title: 'Status', type: 'Choice', choices: ['Active', 'Disabled', 'Completed'], defaultValue: 'Active', indexed: true },
      { name: 'StartDate', title: 'Start date', type: 'DateOnly' },
      { name: 'EndDate', title: 'End date', type: 'DateOnly' },
      { name: 'PlannedHours', title: 'Planned hours', type: 'Number', defaultValue: '0' },
      { name: 'HourlyCost', title: 'Flat hourly cost', type: 'Number', defaultValue: '0' },
      { name: 'BillRate', title: 'Hourly bill rate', type: 'Number', defaultValue: '0' },
      { name: 'MemberRates', title: 'Person-specific rates (JSON)', type: 'Note' },
      { name: 'ColorHex', title: 'Colour', type: 'Text' },
      { name: 'CustomValues', title: 'Custom columns (JSON)', type: 'Note' },
      { name: 'EndNotices', title: 'End-date notices sent', type: 'Text' },
      { name: 'TmProjectId', title: 'Task Manager project ID', type: 'Number', defaultValue: '0', indexed: true }
    ]
  },
  {
    key: 'tasks',
    suffix: 'Tasks',
    description: 'Tasks within projects, with assignees, planned hours and Planner links.',
    titleDisplayName: 'Task',
    security: 'shared',
    fields: [
      { name: 'ProjectId', title: 'Project ID', type: 'Number', indexed: true, required: true },
      { name: 'Details', title: 'Description', type: 'Note' },
      { name: 'AssigneeEmails', title: 'Assignees (emails)', type: 'Note' },
      { name: 'IsBillable', title: 'Billable', type: 'Boolean', defaultValue: '1' },
      { name: 'PlannedHours', title: 'Planned hours', type: 'Number', defaultValue: '0' },
      { name: 'TaskStatus', title: 'Status', type: 'Choice', choices: ['Open', 'In progress', 'Done'], defaultValue: 'Open' },
      { name: 'DueDate', title: 'Due date', type: 'DateOnly' },
      { name: 'SelfAssign', title: 'Self-assign allowed', type: 'Boolean', defaultValue: '1' },
      { name: 'PlannerTaskId', title: 'Planner task ID', type: 'Text', indexed: true },
      { name: 'PlannerPlanId', title: 'Planner plan ID', type: 'Text' },
      { name: 'TmTaskId', title: 'Task Manager task ID', type: 'Number', defaultValue: '0', indexed: true }
    ]
  },
  {
    key: 'activities',
    suffix: 'Activities',
    description: 'Kinds of work tracked on timesheet lines (development, meetings, travel…).',
    titleDisplayName: 'Activity',
    security: 'config',
    fields: [
      { name: 'IsBillable', title: 'Billable', type: 'Boolean', defaultValue: '1' },
      { name: 'IsActive', title: 'Active', type: 'Boolean', defaultValue: '1' },
      { name: 'ColorHex', title: 'Colour', type: 'Text' },
      { name: 'IsTimeOff', title: 'Time off', type: 'Boolean', defaultValue: '0' }
    ]
  },
  {
    key: 'timesheets',
    suffix: 'Timesheets',
    description: 'One row per person per week: status, totals, approval and comments.',
    titleDisplayName: 'Week key',
    versioning: true,
    security: 'timesheet',
    fields: [
      { name: 'UserEmail', title: 'Person email', type: 'Text', indexed: true, required: true },
      { name: 'UserName', title: 'Person', type: 'Text' },
      { name: 'WeekStart', title: 'Week starting', type: 'Text', indexed: true, required: true, description: 'yyyy-mm-dd' },
      { name: 'SheetStatus', title: 'Status', type: 'Choice', choices: ['Draft', 'Submitted', 'Approved', 'Rejected'], defaultValue: 'Draft', indexed: true },
      { name: 'TotalHours', title: 'Total hours', type: 'Number', defaultValue: '0' },
      { name: 'BillableHours', title: 'Billable hours', type: 'Number', defaultValue: '0' },
      { name: 'SubmittedAt', title: 'Submitted', type: 'DateTime' },
      { name: 'DecidedBy', title: 'Approved / rejected by', type: 'Text' },
      { name: 'DecidedAt', title: 'Approved / rejected', type: 'DateTime' },
      { name: 'AutoApproved', title: 'Auto-approved', type: 'Boolean', defaultValue: '0' },
      { name: 'Comments', title: 'Comments (JSON)', type: 'Note' }
    ]
  },
  {
    key: 'entries',
    suffix: 'TimeEntries',
    description: 'Timesheet lines: project, task, activity, billable and hours per day.',
    titleDisplayName: 'Line',
    security: 'timesheet',
    fields: [
      { name: 'SheetId', title: 'Timesheet ID', type: 'Number', indexed: true },
      { name: 'UserEmail', title: 'Person email', type: 'Text', indexed: true },
      { name: 'WeekStart', title: 'Week starting', type: 'Text', indexed: true },
      { name: 'ProjectId', title: 'Project ID', type: 'Number', indexed: true },
      { name: 'TaskId', title: 'Task ID', type: 'Number', defaultValue: '0' },
      { name: 'ActivityId', title: 'Activity ID', type: 'Number', defaultValue: '0' },
      { name: 'IsBillable', title: 'Billable', type: 'Boolean', defaultValue: '1' },
      { name: 'DayHours', title: 'Hours per day (JSON)', type: 'Note' },
      { name: 'DayNotes', title: 'Notes per day (JSON)', type: 'Note' },
      { name: 'TotalHours', title: 'Total hours', type: 'Number', defaultValue: '0' },
      { name: 'CustomValues', title: 'Custom columns (JSON)', type: 'Note' },
      { name: 'CostRate', title: 'Hourly cost (at submit)', type: 'Number', defaultValue: '0' },
      { name: 'BillRate', title: 'Bill rate (at submit)', type: 'Number', defaultValue: '0' },
      { name: 'TmSync', title: 'Task Manager sync (JSON)', type: 'Note' }
    ]
  },
  {
    key: 'tickets',
    suffix: 'SupportTickets',
    description: 'Support requests sent to Brillienta.',
    titleDisplayName: 'Subject',
    security: 'shared',
    fields: [
      { name: 'Details', title: 'Details', type: 'Note' },
      { name: 'Category', title: 'Category', type: 'Text' },
      { name: 'Priority', title: 'Priority', type: 'Choice', choices: ['Low', 'Normal', 'High', 'Urgent'], defaultValue: 'Normal' },
      { name: 'TicketStatus', title: 'Status', type: 'Choice', choices: ['Open', 'In progress', 'Answered', 'Closed'], defaultValue: 'Open' },
      { name: 'RaisedByEmail', title: 'Raised by (email)', type: 'Text' },
      { name: 'Reply', title: 'Reply', type: 'Note' }
    ]
  },
  {
    key: 'notifications',
    suffix: 'Notifications',
    description: 'In-app notifications (the bell).',
    titleDisplayName: 'Title',
    security: 'shared',
    fields: [
      { name: 'Body', title: 'Message', type: 'Note' },
      { name: 'LinkPath', title: 'Link', type: 'Text' },
      { name: 'Recipient', title: 'Recipient email', type: 'Text', indexed: true },
      { name: 'IsRead', title: 'Read', type: 'Boolean', defaultValue: '0' },
      { name: 'Kind', title: 'Kind', type: 'Text' }
    ]
  },
  {
    key: 'emailQueue',
    suffix: 'EmailQueue',
    description: 'Emails waiting to be sent by the Power Automate flow (when email is sent through Power Automate).',
    titleDisplayName: 'Subject',
    security: 'shared',
    fields: [
      { name: 'MailTo', title: 'To', type: 'Note', description: 'Recipients, separated by semicolons.' },
      { name: 'MailBody', title: 'Body (HTML)', type: 'Note' },
      { name: 'MailReplyTo', title: 'Reply to', type: 'Text' },
      { name: 'TemplateKey', title: 'Template', type: 'Text' },
      { name: 'MailStatus', title: 'Status', type: 'Choice', choices: ['Queued', 'Sent', 'Failed'], defaultValue: 'Queued', indexed: true },
      { name: 'MailError', title: 'Error', type: 'Note' },
      { name: 'Kind', title: 'Kind', type: 'Text' }
    ]
  },
  {
    key: 'audit',
    suffix: 'AuditLog',
    description: 'Who changed what: settings, projects, approvals.',
    titleDisplayName: 'Action',
    security: 'adminOnly',
    fields: [
      { name: 'Details', title: 'Details', type: 'Note' },
      { name: 'ByName', title: 'By', type: 'Text' }
    ]
  },
  {
    key: 'reports',
    suffix: 'SavedReports',
    description: 'Custom reports and dashboards.',
    titleDisplayName: 'Report',
    security: 'shared',
    fields: [
      { name: 'OwnerEmail', title: 'Owner email', type: 'Text', indexed: true },
      { name: 'IsShared', title: 'Shared with everyone', type: 'Boolean', defaultValue: '0' },
      { name: 'ReportKind', title: 'Kind', type: 'Choice', choices: ['report', 'dashboard'], defaultValue: 'report' },
      { name: 'Definition', title: 'Definition (JSON)', type: 'Note' }
    ]
  }
];

export function listDef(key: ListKey): IListDef {
  const def = LISTS.filter((l) => l.key === key)[0];
  if (!def) throw new Error('Unknown list ' + key);
  return def;
}

export function listTitle(prefix: string, key: ListKey): string {
  return prefix + listDef(key).suffix;
}

function xmlEscape(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** CAML field definition passed to fields/createfieldasxml. */
export function fieldSchemaXml(f: IFieldDef): string {
  const common =
    `Name="${f.name}" StaticName="${f.name}" DisplayName="${xmlEscape(f.title)}"` +
    ` Required="${f.required ? 'TRUE' : 'FALSE'}"` +
    (f.indexed ? ' Indexed="TRUE"' : '') +
    (f.description ? ` Description="${xmlEscape(f.description)}"` : '');
  const def = f.defaultValue !== undefined ? `<Default>${xmlEscape(f.defaultValue)}</Default>` : '';
  const choices = (f.choices || []).map((c) => `<CHOICE>${xmlEscape(c)}</CHOICE>`).join('');
  switch (f.type) {
    case 'Text':
      return `<Field Type="Text" ${common} MaxLength="255">${def}</Field>`;
    case 'Note':
      return `<Field Type="Note" ${common} NumLines="6" RichText="FALSE" AppendOnly="FALSE" UnlimitedLengthInDocumentLibrary="TRUE" />`;
    case 'Number':
      return `<Field Type="Number" ${common}>${def}</Field>`;
    case 'Boolean':
      return `<Field Type="Boolean" ${common}>${def || '<Default>0</Default>'}</Field>`;
    case 'DateTime':
      return `<Field Type="DateTime" ${common} Format="DateTime" />`;
    case 'DateOnly':
      return `<Field Type="DateTime" ${common} Format="DateOnly" />`;
    case 'Choice':
      return `<Field Type="Choice" ${common} Format="Dropdown" FillInChoice="TRUE"><CHOICES>${choices}</CHOICES>${def}</Field>`;
    default:
      throw new Error('Unsupported field type ' + (f.type as string));
  }
}

/** The TypeAsString SharePoint reports for a field of this definition. */
export function expectedTypeAsString(f: IFieldDef): string {
  return f.type === 'DateOnly' ? 'DateTime' : f.type;
}

/** $select for reading items of a list. */
export function selectFor(key: ListKey): string {
  return ['Id', 'Title', 'Created', 'Modified', 'Author/Title', 'Author/EMail'].concat(listDef(key).fields.map((f) => f.name)).join(',');
}

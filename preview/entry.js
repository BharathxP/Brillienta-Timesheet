// Local preview: renders the real app (compiled lib/) against an in-memory SharePoint with a sample organisation.
//   ?role=admin (default) | pm | manager | member   who is signed in
//   ?plan=standard | pro                            run as that plan (default: the Enterprise trial)
//   ?fresh=1                                        a site where the app is not installed yet (setup wizard)
//   ?installed=30                                   installed 30 days ago (the trial has ended)
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { initializeIcons } from '@fluentui/react';
import App from '../lib/webparts/brillientaTimesheet/components/App';
import { defaultConfig, DEFAULT_ACTIVITIES } from '../lib/webparts/brillientaTimesheet/schema/defaults';
import { SCHEMA_VERSION } from '../lib/webparts/brillientaTimesheet/schema/schema';
import { addDays, today, weekStartOf, weekDays } from '../lib/webparts/brillientaTimesheet/services/dates';
import { createFakeHttp, fakeLists, seed } from './fakeSharePoint';
import { version as APP_VERSION } from '../package.json';
import { PREVIEW_PRIVATE_JWK } from './previewKey';

const params = new URLSearchParams(window.location.search);
const role = params.get('role') || 'admin';
const WEB = 'https://contoso.sharepoint.com/sites/timesheet';
const PEOPLE = [
  { name: 'Priya Nair', email: 'priya.nair@contoso.com', role: 'admin', manager: '', dept: 'Operations', title: 'Head of Delivery', cost: 2200, bill: 4000 },
  { name: 'Vikram Rao', email: 'vikram.rao@contoso.com', role: 'programManager', manager: 'priya.nair@contoso.com', dept: 'Delivery', title: 'Programme Manager', cost: 1900, bill: 3600 },
  { name: 'Arjun Mehta', email: 'arjun.mehta@contoso.com', role: 'projectManager', manager: 'vikram.rao@contoso.com', dept: 'Delivery', title: 'Project Manager', cost: 1600, bill: 3200 },
  { name: 'Meera Iyer', email: 'meera.iyer@contoso.com', role: 'teamManager', manager: 'priya.nair@contoso.com', dept: 'Engineering', title: 'Engineering Manager', cost: 1800, bill: 3400 },
  { name: 'Rahul Sharma', email: 'rahul.sharma@contoso.com', role: 'member', manager: 'meera.iyer@contoso.com', dept: 'Engineering', title: 'Senior Developer', cost: 1400, bill: 2800 },
  { name: 'Sara Khan', email: 'sara.khan@contoso.com', role: 'member', manager: 'meera.iyer@contoso.com', dept: 'Engineering', title: 'UX Designer', cost: 1300, bill: 2600 },
  { name: 'Tom Becker', email: 'tom.becker@contoso.com', role: 'member', manager: 'meera.iyer@contoso.com', dept: 'Engineering', title: 'QA Engineer', cost: 1100, bill: 2200 },
  { name: 'Ana Lopez (Fabrikam)', email: 'ana.lopez@fabrikam.com', role: 'member', manager: 'arjun.mehta@contoso.com', dept: 'Contractors', title: 'Consultant', cost: 2000, bill: 3000, external: true }
];
const SIGNED_IN = { admin: 0, pgm: 1, pm: 2, manager: 3, member: 4 }[role] || 0;
const me = PEOPLE[SIGNED_IN];
const user = { Id: 7, Title: me.name, Email: me.email, LoginName: 'i:0#.f|membership|' + me.email, IsSiteAdmin: role === 'admin' };
const author = { Id: user.Id, Title: user.Title, EMail: user.Email };
const http = createFakeHttp(user);
window.__tsLists = fakeLists;

function render() {
  ReactDom.render(
    React.createElement(App, {
      http,
      webUrl: WEB,
      siteTitle: 'Contoso Timesheet',
      user: { displayName: me.name, email: me.email, loginName: user.LoginName },
      permissions: { manageLists: role === 'admin', manageWeb: role === 'admin' },
      tenantId: '6f1c2a9e-0b7d-4c55-9e3a-2d8f41b07c13',
      prefix: 'TS_',
      appVersion: APP_VERSION + '.0',
      isEditMode: false
    }),
    document.getElementById('root')
  );
}

async function previewKey(plan, tenant) {
  const b64url = (bytes) => btoa(String.fromCharCode.apply(null, Array.from(bytes))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  const payload = { v: 1, product: 'timesheet', id: 'BT-PREVIEW', plan, tenant, customer: 'Contoso (preview)', users: 0, issued: today(), accountManager: { name: 'Kavya Reddy', email: 'kavya@brillienta.com', phone: '+91 80 0000 0000' } };
  const signed = 'BT1.' + b64url(new TextEncoder().encode(JSON.stringify(payload)));
  const key = await crypto.subtle.importKey('jwk', PREVIEW_PRIVATE_JWK, { name: 'ECDSA', namedCurve: 'P-256' }, false, ['sign']);
  const sig = new Uint8Array(await crypto.subtle.sign({ name: 'ECDSA', hash: 'SHA-256' }, key, new TextEncoder().encode(signed)));
  return signed + '.' + b64url(sig);
}

/** A small Brillienta Task Manager installation (TM_ lists), as its own app writes them. */
function seedTaskManager(author) {
  const t = today();
  const who = (e) => ({ EMail: e });
  seed('TM_Projects', [
    { Title: 'Customer portal', ProjectCode: 'PORTAL', ProjectStatus: 'Active', PlannedStart: addDays(t, -14) + 'T00:00:00Z', PlannedEnd: addDays(t, 75) + 'T00:00:00Z', BudgetHours: 240, ProjectOwner: who('arjun.mehta@contoso.com'), ProjectMembers: [who('rahul.sharma@contoso.com'), who('sara.khan@contoso.com')], Details: '<p>Self-service portal for Contoso customers.</p>', ColorHex: '#0e7490' },
    { Title: 'Mobile app v2', ProjectCode: 'APP-02', ProjectStatus: 'Active', PlannedStart: addDays(t, -30) + 'T00:00:00Z', PlannedEnd: addDays(t, 120) + 'T00:00:00Z', BudgetHours: 950, ProjectOwner: who('arjun.mehta@contoso.com'), ProjectMembers: [] },
    { Title: 'Old intranet', ProjectCode: 'INTRA', ProjectStatus: 'Cancelled', BudgetHours: 0, ProjectOwner: who('priya.nair@contoso.com'), ProjectMembers: [] }
  ], author);
  const pid = (code) => fakeLists().TM_Projects.items.filter((x) => x.ProjectCode === code)[0].Id;
  seed('TM_Tasks', [
    { Title: 'Sign-in and accounts', TaskNumber: 'TM-101', ProjectId: pid('PORTAL'), StatusKind: 'InProgress', AssigneeEmail: 'rahul.sharma@contoso.com', DueDate: addDays(t, 10) + 'T00:00:00Z', EstimateHours: 24, SpentMinutes: 120 },
    { Title: 'Order history', TaskNumber: 'TM-102', ProjectId: pid('PORTAL'), StatusKind: 'NotStarted', AssigneeEmail: 'sara.khan@contoso.com', DueDate: addDays(t, 20) + 'T00:00:00Z', EstimateHours: 16, SpentMinutes: 0 },
    { Title: 'Push notifications', TaskNumber: 'TM-201', ProjectId: pid('APP-02'), StatusKind: 'Waiting', AssigneeEmail: 'rahul.sharma@contoso.com', DueDate: addDays(t, 5) + 'T00:00:00Z', EstimateHours: 12, SpentMinutes: 0 }
  ], author);
  const tid = (n) => fakeLists().TM_Tasks.items.filter((x) => x.TaskNumber === n)[0].Id;
  const week = weekStartOf(t, 1);
  // Rahul logged two hours on Monday in Task Manager's task timer.
  seed('TM_TimeEntries', [{ Title: 'Timer', TaskId: tid('TM-101'), ProjectId: pid('PORTAL'), DurationMinutes: 120, WorkNote: 'OAuth setup', Billable: true, WorkDate: week }], { Id: 104, Title: 'Rahul Sharma', EMail: 'rahul.sharma@contoso.com' });
  seed('TM_TimeEntries', [{ Title: 'Timer', TaskId: tid('TM-201'), ProjectId: pid('APP-02'), DurationMinutes: 90, WorkNote: 'Design review', Billable: true, WorkDate: addDays(week, 1) }], { Id: 101, Title: 'Priya Nair', EMail: 'priya.nair@contoso.com' });
}

async function main() {
  initializeIcons(undefined, { disableWarnings: true });
  if (params.get('fresh')) return render();
  const cfg = defaultConfig('Contoso');
  cfg.orgName = 'Contoso';
  cfg.limits = { dailyMin: 0, dailyMax: 10, weeklyMin: 40, weeklyMax: 50, enforcement: 'warn' };
  cfg.autoApproval = { enabled: true, roles: ['admin', 'programManager'], requireWithinRules: true, maxHours: 0 };
  cfg.customColumns = [{ key: 'ticket', label: 'Ticket #', type: 'text', level: 'entry', choices: [], required: false, inReports: true }, { key: 'po', label: 'PO number', type: 'text', level: 'project', choices: [], required: false, inReports: false }];
  // ?tm=1: connected to Brillienta Task Manager (TM_ lists on the same site, seeded below).
  // ?leave=1: Time off Manager 365 connected (a leave list on the same site, seeded below).
  if (params.get('leave')) cfg.timeOff = { ...cfg.timeOff, enabled: true, autoFill: true };
  if (params.get('tm')) cfg.taskManager = { ...cfg.taskManager, enabled: true, autoSync: true, pushTime: 'approved', importTime: true };
  cfg.holidays = [{ date: addDays(weekStartOf(today(), 1), 14 + 2), name: 'Company day' }];
  const installedAt = params.get('installed') ? new Date(Date.now() - Number(params.get('installed')) * 86400000).toISOString() : new Date(Date.now() - 3 * 86400000).toISOString();
  seed(
    'TS_Settings',
    [
      { Title: 'SchemaVersion', SettingValue: String(SCHEMA_VERSION) },
      { Title: 'AppVersion', SettingValue: APP_VERSION + '.0' },
      { Title: 'Config', SettingValue: JSON.stringify(cfg) },
      { Title: 'InstalledAt', SettingValue: installedAt },
      // Reminders already ran (the preview does not send them on open).
      { Title: 'ReminderRun', SettingValue: new Date().toISOString() + '|preview' }
    ],
    author
  );
  seed('TS_Users', PEOPLE.map((p) => ({ Title: p.name, Email: p.email, UserRole: p.role, ManagerEmail: p.manager, Department: p.dept, JobTitle: p.title, CostRate: p.cost, BillRate: p.bill, IsExternal: !!p.external, IsActive: true, Created: '2026-01-05T09:00:00Z' })), author);
  seed('TS_Programs', [{ Title: 'Digital transformation', Details: 'Client-facing digital programmes.', ManagerEmails: 'vikram.rao@contoso.com', ColorHex: '#8a5cd1' }], author);
  const t = today();
  const prog = fakeLists().TS_Programs.items[0].Id;
  const P = [
    { Title: 'Website redesign', ProjectCode: 'WEB-01', Client: 'Contoso Retail', ProgramId: prog, ManagerEmails: 'arjun.mehta@contoso.com', IsBillable: true, ProjectStatus: 'Active', StartDate: addDays(t, -60), EndDate: addDays(t, 6), PlannedHours: 420, HourlyCost: 1500, BillRate: 3000, ColorHex: '#2a78d6', CustomValues: JSON.stringify({ po: 'PO-7781' }) },
    { Title: 'Mobile app', ProjectCode: 'APP-02', Client: 'Fabrikam', ProgramId: prog, ManagerEmails: 'arjun.mehta@contoso.com', IsBillable: true, ProjectStatus: 'Active', StartDate: addDays(t, -30), EndDate: addDays(t, 120), PlannedHours: 900, HourlyCost: 1500, BillRate: 3200, ColorHex: '#1baf7a' },
    { Title: 'Data platform', ProjectCode: 'DATA-03', Client: 'Northwind', ProgramId: 0, ManagerEmails: 'meera.iyer@contoso.com', IsBillable: true, ProjectStatus: 'Active', StartDate: addDays(t, -90), EndDate: addDays(t, 60), PlannedHours: 300, HourlyCost: 1600, BillRate: 3400, ColorHex: '#d18a2a' },
    { Title: 'Internal', ProjectCode: 'INT', Client: '', ProgramId: 0, ManagerEmails: 'priya.nair@contoso.com', IsBillable: false, ProjectStatus: 'Active', StartDate: '', EndDate: '', PlannedHours: 0, HourlyCost: 0, BillRate: 0, ColorHex: '#17232e' },
    { Title: 'Legacy CRM support', ProjectCode: 'CRM-00', Client: 'Contoso Retail', ProgramId: 0, ManagerEmails: 'arjun.mehta@contoso.com', IsBillable: true, ProjectStatus: 'Completed', StartDate: addDays(t, -300), EndDate: addDays(t, -40), PlannedHours: 200, HourlyCost: 1400, BillRate: 2800, ColorHex: '#94a3b8' }
  ];
  seed('TS_Projects', P, author);
  const pid = (code) => fakeLists().TS_Projects.items.filter((x) => x.ProjectCode === code)[0].Id;
  seed(
    'TS_Tasks',
    [
      ['WEB-01', 'Discovery workshops', 'Done', ''],
      ['WEB-01', 'Visual design', 'In progress', 'sara.khan@contoso.com'],
      ['WEB-01', 'Front-end build', 'In progress', 'rahul.sharma@contoso.com'],
      ['WEB-01', 'Accessibility review', 'Open', ''],
      ['APP-02', 'API integration', 'Open', 'rahul.sharma@contoso.com'],
      ['APP-02', 'Test automation', 'Open', 'tom.becker@contoso.com'],
      ['APP-02', 'App store release', 'Open', ''],
      ['DATA-03', 'Pipeline design', 'In progress', 'meera.iyer@contoso.com'],
      ['INT', 'Hiring interviews', 'Open', '']
    ].map(([code, title, status, who], i) => ({ Title: title, ProjectId: pid(code), TaskStatus: status, AssigneeEmails: who, IsBillable: code !== 'INT', PlannedHours: 60, SelfAssign: true, DueDate: addDays(t, 7 * (i % 5) - 3) })),
    author
  );
  seed('TS_Activities', DEFAULT_ACTIVITIES.map((a) => ({ Title: a.title, IsBillable: a.billable, IsActive: true, ColorHex: a.color, IsTimeOff: !!a.timeOff })), author);
  const task = (title) => fakeLists().TS_Tasks.items.filter((x) => x.Title === title)[0].Id;
  const act = (title) => fakeLists().TS_Activities.items.filter((x) => x.Title === title)[0].Id;
  // Four past weeks for everyone: approved, then submitted (waiting), the current week a draft.
  const thisWeek = weekStartOf(t, 1);
  const plan = {
    'priya.nair@contoso.com': [['INT', '', 'Project management', 0, 5]],
    'vikram.rao@contoso.com': [['WEB-01', '', 'Project management', 0, 3], ['APP-02', '', 'Meetings', 0, 4]],
    'arjun.mehta@contoso.com': [['WEB-01', 'Front-end build', 'Project management', 0, 4], ['APP-02', 'API integration', 'Meetings', 0, 3.5]],
    'meera.iyer@contoso.com': [['DATA-03', 'Pipeline design', 'Development', 0, 6], ['INT', 'Hiring interviews', 'Internal', 0, 2]],
    'rahul.sharma@contoso.com': [['WEB-01', 'Front-end build', 'Development', 'FE-142', 6], ['APP-02', 'API integration', 'Development', 'API-9', 2]],
    'sara.khan@contoso.com': [['WEB-01', 'Visual design', 'Design', '', 7], ['INT', '', 'Training', '', 1]],
    'tom.becker@contoso.com': [['APP-02', 'Test automation', 'Testing', '', 8]],
    'ana.lopez@fabrikam.com': [['APP-02', '', 'Development', '', 8]]
  };
  const sheets = [];
  const entries = [];
  PEOPLE.forEach((p, n) => {
    [4, 3, 2, 1, 0].forEach((ago) => {
      const w = addDays(thisWeek, -7 * ago);
      if (p.email === me.email && ago === 0) return; // my current week starts empty (or from the sample below)
      const status = ago >= 2 ? 'Approved' : ago === 1 ? (n % 3 === 0 ? 'Draft' : 'Submitted') : 'Draft';
      const lines = plan[p.email].map(([code, tk, a, ticket, h]) => {
        const hours = {};
        weekDays(w)
          .slice(0, 5)
          .forEach((d, i) => {
            if (ago === 0 && d > t) return;
            hours[d] = Math.max(0, h + ((i + n + ago) % 3) * 0.5 - 0.5);
          });
        const total = Object.keys(hours).reduce((s, k) => s + hours[k], 0);
        return { code, tk, a, ticket, hours, total };
      });
      const total = lines.reduce((s, l) => s + l.total, 0);
      sheets.push({ Title: `${p.email}|${w}`, UserEmail: p.email, UserName: p.name, WeekStart: w, SheetStatus: status, TotalHours: total, BillableHours: lines.filter((l) => l.code !== 'INT').reduce((s, l) => s + l.total, 0), SubmittedAt: status !== 'Draft' ? addDays(w, 5) + 'T17:00:00Z' : null, DecidedBy: status === 'Approved' ? 'Priya Nair' : '', DecidedAt: status === 'Approved' ? addDays(w, 7) + 'T10:00:00Z' : null, Comments: JSON.stringify(status === 'Approved' ? [{ at: addDays(w, 5) + 'T17:00:00Z', by: p.name, byEmail: p.email, text: '', action: 'submitted' }, { at: addDays(w, 7) + 'T10:00:00Z', by: 'Priya Nair', byEmail: 'priya.nair@contoso.com', text: '', action: 'approved' }] : []) });
      lines.forEach((l) => entries.push({ __sheet: `${p.email}|${w}`, UserEmail: p.email, WeekStart: w, ProjectId: pid(l.code), TaskId: l.tk ? task(l.tk) : 0, ActivityId: act(l.a), IsBillable: l.code !== 'INT', DayHours: JSON.stringify(l.hours), DayNotes: '{}', TotalHours: l.total, CustomValues: JSON.stringify(l.ticket ? { ticket: l.ticket } : {}), CostRate: status === 'Approved' ? p.cost : 0, BillRate: status === 'Approved' ? p.bill : 0 }));
    });
  });
  seed('TS_Timesheets', sheets, author);
  const sheetId = (k) => fakeLists().TS_Timesheets.items.filter((x) => x.Title === k)[0].Id;
  seed('TS_TimeEntries', entries.map((e) => { const { __sheet, ...rest } = e; return { ...rest, SheetId: sheetId(__sheet), Title: __sheet }; }), author);
  seed('TS_Notifications', [{ Title: 'Timesheet to approve: Rahul Sharma, week of last week', Body: 'Rahul Sharma submitted 40 hours (32 billable).', LinkPath: 'approvals', Recipient: me.email, IsRead: false, Kind: 'sheetSubmitted' }], author);
  if (params.get('tm')) seedTaskManager(author);
  if (params.get('leave')) {
    // Approved leave for the signed-in person: Wednesday to the following Monday (it spans a weekend), and a rejected request.
    const w = weekStartOf(today(), 1);
    seed('TOM_LeaveRequests', [
      { Title: 'Family trip', EmployeeEmail: me.email, StartDate: addDays(w, 2) + 'T00:00:00Z', EndDate: addDays(w, 7) + 'T00:00:00Z', RequestStatus: 'Approved', LeaveType: 'Annual leave' },
      { Title: 'Rejected', EmployeeEmail: me.email, StartDate: addDays(w, 0) + 'T00:00:00Z', EndDate: addDays(w, 0) + 'T00:00:00Z', RequestStatus: 'Rejected', LeaveType: 'Sick leave' }
    ], author);
  }
  ['TS_Tickets', 'TS_SupportTickets', 'TS_EmailQueue', 'TS_AuditLog', 'TS_SavedReports'].forEach((l) => seed(l, [], author));
  // ?plan=: a licence for that plan, as if activated (signed with the preview key pair).
  if (params.get('plan')) seed('TS_Settings', [{ Title: 'LicenseKey', SettingValue: await previewKey(params.get('plan'), '6f1c2a9e-0b7d-4c55-9e3a-2d8f41b07c13') }], author);
  render();
}

main().catch((e) => {
  document.body.textContent = 'Preview failed: ' + e.message;
});

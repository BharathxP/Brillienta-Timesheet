// SPFx libraries expect a browser global when they load (the services that import them are not called here).
globalThis.window = globalThis.window || globalThis;

// Shared fixtures. Tests run against the CommonJS build (npm test builds first).
const lib = (p) => require('../lib-commonjs/webparts/brillientaTimesheet/' + p);

let n = 0;
function user(o = {}) {
  n++;
  return { id: n, name: `Person ${n}`, email: `p${n}@contoso.com`, role: 'member', managerEmail: '', department: '', jobTitle: '', costRate: 0, billRate: 0, external: false, active: true, ruleId: '', workDays: [], ...o };
}
function project(o = {}) {
  n++;
  return { id: n, title: `Project ${n}`, code: '', client: '', description: '', programId: 0, managerEmails: [], memberEmails: [], billable: true, status: 'Active', startDate: '', endDate: '', plannedHours: 0, hourlyCost: 0, billRate: 0, memberRates: {}, color: '', custom: {}, endNotices: [], tmProjectId: 0, ...o };
}
function entry(o = {}) {
  n++;
  return { id: n, sheetId: 1, userEmail: 'p@contoso.com', weekStart: '2026-11-16', projectId: 1, taskId: 0, activityId: 0, billable: true, hours: {}, notes: {}, total: 0, custom: {}, costRate: 0, billRate: 0, tmSync: {}, ...o };
}
function sheet(o = {}) {
  n++;
  return { id: n, userEmail: 'p@contoso.com', userName: 'P', weekStart: '2026-11-16', status: 'Submitted', totalHours: 0, billableHours: 0, submittedAt: '', decidedBy: '', decidedAt: '', autoApproved: false, comments: [], ...o };
}
module.exports = { lib, user, project, entry, sheet };

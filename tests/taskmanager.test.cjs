const test = require('node:test');
const assert = require('node:assert/strict');
const { lib, project, entry } = require('./helpers.cjs');
const T = lib('services/taskManagerSync.js');

const tmProject = (o = {}) => ({ id: 1, title: 'Portal', code: 'POR', status: 'Active', start: '2026-11-01', end: '2027-01-31', budgetHours: 200, ownerEmail: 'own@c.com', memberEmails: [], description: '', color: '', ...o });
const tmTask = (o = {}) => ({ id: 11, title: 'Login page', number: 'TM-11', projectId: 1, statusKind: 'InProgress', assigneeEmail: 'dev@c.com', due: '2026-12-01', estimateHours: 16, spentMinutes: 0, ...o });
const tsTask = (o = {}) => ({ id: 500, title: 'x', projectId: 100, description: '', assigneeEmails: [], billable: true, plannedHours: 0, status: 'Open', dueDate: '', selfAssign: true, plannerTaskId: '', plannerPlanId: '', tmTaskId: 0, ...o });

test('Task Manager rows are read as Task Manager writes them', () => {
  const p = T.toTmProject({ Id: 3, Title: 'Portal', ProjectCode: 'POR', ProjectStatus: 'OnHold', PlannedStart: '2026-11-01T00:00:00Z', ProjectOwner: { EMail: 'own@c.com' }, ProjectMembers: [{ EMail: 'a@c.com' }, { EMail: 'b@c.com' }], Details: '<p>Client <b>portal</b></p>' });
  assert.equal(p.ownerEmail, 'own@c.com');
  assert.deepEqual(p.memberEmails, ['a@c.com', 'b@c.com']);
  assert.equal(p.description, 'Client portal');
  assert.equal(T.toTmTimeEntry({ Id: 1, TaskId: 2, DurationMinutes: 90, WorkDate: '2026-11-17', Billable: true, Author: { EMail: 'x@c.com' } }).authorEmail, 'x@c.com');
  assert.equal(T.taskStatusFromTm('Waiting'), 'In progress');
  assert.equal(T.taskStatusFromTm('Cancelled'), 'Done');
  assert.equal(T.projectStatusFromTm('OnHold'), 'Disabled');
  assert.equal(T.projectStatusFromTm('Planning'), 'Active');
});

test('projects: create, adopt by code, update only what changed, respect the status filter', () => {
  const existing = project({ id: 100, code: 'por', title: 'Old name', hourlyCost: 900, billRate: 2000 });
  const plan = T.planProjects([tmProject(), tmProject({ id: 2, code: 'NEW', title: 'New one', memberEmails: ['m@c.com'] }), tmProject({ id: 3, code: 'X', status: 'Cancelled' })], [existing], ['Active', 'Planning'], ['#111']);
  assert.equal(plan.update.length, 1);
  assert.equal(plan.update[0].id, 100);
  assert.equal(plan.update[0].patch.title, 'Portal');
  assert.equal(plan.update[0].patch.tmProjectId, 1);
  assert.equal(plan.update[0].patch.hourlyCost, undefined); // rates stay as set in Timesheet
  assert.equal(plan.create.length, 1);
  assert.deepEqual(plan.create[0].memberEmails, ['m@c.com', 'own@c.com']);
  assert.equal(plan.create[0].tmProjectId, 2);
  // already in sync: nothing to write
  const synced = project({ id: 100, code: 'POR', title: 'Portal', status: 'Active', startDate: '2026-11-01', endDate: '2027-01-31', plannedHours: 200, managerEmails: ['own@c.com'], memberEmails: [], tmProjectId: 1 });
  assert.equal(T.planProjects([tmProject()], [synced], [], []).update.length, 0);
});

test('tasks of linked projects: create open ones, update linked ones', () => {
  const projects = [project({ id: 100, tmProjectId: 1, billable: false })];
  const plan = T.planTasks([tmTask(), tmTask({ id: 12, statusKind: 'Done' }), tmTask({ id: 13, projectId: 99 })], [], projects);
  assert.equal(plan.create.length, 1);
  assert.equal(plan.create[0].title, 'Login page (TM-11)');
  assert.deepEqual(plan.create[0].assigneeEmails, ['dev@c.com']);
  assert.equal(plan.create[0].billable, false);
  assert.equal(plan.create[0].status, 'In progress');
  const upd = T.planTasks([tmTask({ statusKind: 'Done' })], [tsTask({ tmTaskId: 11, projectId: 100, title: 'Login page (TM-11)', status: 'In progress', dueDate: '2026-12-01', plannedHours: 16, assigneeEmails: ['dev@c.com'] })], projects);
  assert.deepEqual(upd.update, [{ id: 500, patch: { status: 'Done' } }]);
});

test('importing logged time: grouped per task, notes joined, skipped when not linked or already imported', () => {
  const tasks = [tsTask({ id: 500, projectId: 100, tmTaskId: 11 })];
  const projects = [project({ id: 100, tmProjectId: 1, billable: true })];
  const entries = [
    { id: 1, taskId: 11, projectId: 1, minutes: 90, note: 'a', billable: true, workDate: '2026-11-16', authorEmail: 'me@c.com' },
    { id: 2, taskId: 11, projectId: 1, minutes: 30, note: 'b', billable: true, workDate: '2026-11-16', authorEmail: 'me@c.com' },
    { id: 3, taskId: 77, projectId: 9, minutes: 60, note: '', billable: true, workDate: '2026-11-17', authorEmail: 'me@c.com' },
    { id: 4, taskId: 11, projectId: 1, minutes: 60, note: '', billable: true, workDate: '2026-11-30', authorEmail: 'me@c.com' },
    { id: 5, taskId: 11, projectId: 1, minutes: 60, note: '', billable: true, workDate: '2026-11-18', authorEmail: 'me@c.com' }
  ];
  const r = T.importLines(entries, '2026-11-16', 'me@c.com', tasks, projects, [5]);
  assert.equal(r.lines.length, 1);
  assert.equal(r.skipped, 1);
  assert.deepEqual(r.lines[0].hours, { '2026-11-16': 2 });
  assert.equal(r.lines[0].notes['2026-11-16'], 'a; b');
  assert.deepEqual(r.lines[0].tmSync.imported, [1, 2]);
});

test('writing back: linked task lines only, one item per day, imported lines never', () => {
  const tasks = [tsTask({ id: 500, tmTaskId: 11 }), tsTask({ id: 501 })];
  const line = entry({ taskId: 500, hours: { '2026-11-17': 1.5, '2026-11-16': 8, '2026-11-18': 0 }, notes: { '2026-11-16': 'build' } });
  assert.deepEqual(T.pushItems(line, tasks), { tmTaskId: 11, items: [{ workDate: '2026-11-16', minutes: 480, note: 'build' }, { workDate: '2026-11-17', minutes: 90, note: '' }] });
  assert.equal(T.pushItems(entry({ taskId: 501, hours: { '2026-11-16': 2 } }), tasks).tmTaskId, 0);
  assert.equal(T.pushItems({ ...line, tmSync: { imported: [1] } }, tasks).items.length, 0);
});

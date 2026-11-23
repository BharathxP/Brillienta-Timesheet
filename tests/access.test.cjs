const test = require('node:test');
const assert = require('node:assert/strict');
const { lib, user, project, entry, sheet } = require('./helpers.cjs');
const A = lib('services/access.js');
const { defaultConfig } = lib('schema/defaults.js');

const boss = user({ email: 'boss@c.com', role: 'teamManager' });
const dev = user({ email: 'dev@c.com', managerEmail: 'boss@c.com' });
const intern = user({ email: 'intern@c.com', managerEmail: 'dev@c.com' });
const pm = user({ email: 'pm@c.com', role: 'projectManager' });
const pgm = user({ email: 'pgm@c.com', role: 'programManager' });
const p1 = project({ id: 101, managerEmails: ['pm@c.com'], programId: 7 });
const p2 = project({ id: 102, managerEmails: ['other@c.com'], memberEmails: ['dev@c.com'] });
const dir = { users: [boss, dev, intern, pm, pgm], projects: [p1, p2], programs: [{ id: 7, title: 'G', description: '', managerEmails: ['pgm@c.com'], color: '' }], tasks: [] };
const s = sheet({ userEmail: 'intern@c.com' });
const lines = [entry({ projectId: 101, total: 8 })];

test('reporting lines (direct and indirect, loops cut)', () => {
  assert.deepEqual(A.reportsOf(dir.users, 'boss@c.com').map((u) => u.email), ['dev@c.com', 'intern@c.com']);
  assert.deepEqual(A.reportsOf(dir.users, 'boss@c.com', true).map((u) => u.email), ['dev@c.com']);
  const loop = [user({ email: 'a@c.com', managerEmail: 'b@c.com' }), user({ email: 'b@c.com', managerEmail: 'a@c.com' })];
  assert.equal(A.reportsOf(loop, 'a@c.com').length, 1);
});

test('who approves: manager chain, project and programme managers, never yourself', () => {
  const either = defaultConfig();
  assert.equal(A.canApprove('teamManager', 'boss@c.com', s, lines, dir, either), true); // indirect manager
  assert.equal(A.canApprove('projectManager', 'pm@c.com', s, lines, dir, either), true);
  assert.equal(A.canApprove('programManager', 'pgm@c.com', s, lines, dir, either), true);
  assert.equal(A.canApprove('projectManager', 'pm@c.com', s, lines, dir, { ...either, approvalBy: 'manager' }), false);
  assert.equal(A.canApprove('teamManager', 'boss@c.com', s, lines, dir, { ...either, approvalBy: 'project' }), false);
  assert.equal(A.canApprove('admin', 'intern@c.com', s, lines, dir, either), false);
  assert.equal(A.canApprove('member', 'x@c.com', s, lines, dir, either), false);
});

test('approvers asked: manager and project managers, else administrators', () => {
  assert.deepEqual(A.approversFor('intern@c.com', lines, dir, defaultConfig(), ['adm@c.com']), ['dev@c.com', 'pm@c.com']); // either: both asked
  assert.deepEqual(A.approversFor('intern@c.com', lines, dir, { ...defaultConfig(), approvalBy: 'manager' }, ['adm@c.com']), ['dev@c.com']);
  assert.deepEqual(A.approversFor('intern@c.com', lines, dir, { ...defaultConfig(), approvalBy: 'project' }, ['adm@c.com']), ['pm@c.com']);
  assert.deepEqual(A.approversFor('pm@c.com', [entry({ projectId: 101, total: 2 })], dir, { ...defaultConfig(), approvalBy: 'project' }, ['adm@c.com']), ['adm@c.com']);
});

test('roles: site owners are admins; programme manager needs Pro', () => {
  assert.equal(A.effectiveRole(dev, { siteAdmin: true, configAdmin: false, programManagerAllowed: true }), 'admin');
  assert.equal(A.effectiveRole(pgm, { siteAdmin: false, configAdmin: false, programManagerAllowed: false }), 'projectManager');
  assert.equal(A.effectiveRole(undefined, { siteAdmin: false, configAdmin: false, programManagerAllowed: true }), 'member');
});

test('projects and self-assignable tasks a person may use', () => {
  assert.deepEqual(A.projectsFor(dir, 'intern@c.com').map((p) => p.id), [101]);
  assert.deepEqual(A.projectsFor(dir, 'dev@c.com').map((p) => p.id), [101, 102]);
  const d2 = { ...dir, projects: [project({ id: 5, status: 'Disabled' }), project({ id: 6, endDate: '2026-11-01' })] };
  assert.deepEqual(A.projectsFor(d2, 'dev@c.com', '2026-11-16').map((p) => p.id), []);
  const tasks = [{ id: 1, projectId: 101, selfAssign: true, status: 'Open', assigneeEmails: [] }, { id: 2, projectId: 101, selfAssign: false, status: 'Open', assigneeEmails: [] }, { id: 3, projectId: 102, selfAssign: true, status: 'Open', assigneeEmails: [] }];
  assert.deepEqual(A.selfAssignable({ ...dir, tasks }, 'intern@c.com').map((t) => t.id), [1]);
});

test('auto approval by role, rules and hours', () => {
  const cfg = { ...defaultConfig(), autoApproval: { enabled: true, roles: ['projectManager'], requireWithinRules: true, maxHours: 45 } };
  assert.equal(A.shouldAutoApprove(cfg, 'projectManager', 40, []), true);
  assert.equal(A.shouldAutoApprove(cfg, 'member', 40, []), false);
  assert.equal(A.shouldAutoApprove(cfg, 'projectManager', 50, []), false);
  assert.equal(A.shouldAutoApprove(cfg, 'projectManager', 40, [{ level: 'warning', message: 'x' }]), false);
});

test('cost and bill rates', () => {
  const u = user({ email: 'r@c.com', costRate: 1000, billRate: 2000 });
  const p = project({ hourlyCost: 1500, billRate: 2500, memberRates: { 'R@c.com': 3000 } });
  assert.equal(A.costRate('userFirst', u, p, 1), 1000);
  assert.equal(A.costRate('projectFirst', u, p, 1), 1500);
  assert.equal(A.costRate('userFirst', user(), p, 1), 1500);
  assert.equal(A.costRate('projectOnly', u, project(), 7), 7);
  assert.equal(A.billRate(u, p, 1, true), 3000); // person-specific rate on the project
  assert.equal(A.billRate(u, project({ billRate: 2500 }), 1, true), 2000);
  assert.equal(A.billRate(u, p, 1, false), 2500); // without user-specific rates (Standard)
  assert.equal(A.defaultBillable(project({ billable: false }), { billable: true }, true), false);
});

test('security groups: each person in the highest group, from roles and reporting lines', () => {
  const users = [
    user({ email: 'Adm@c.com', role: 'admin' }),
    user({ email: 'pm@c.com', role: 'projectManager' }),
    user({ email: 'boss@c.com' }), // someone reports to them
    user({ email: 'dev@c.com', managerEmail: 'boss@c.com' }),
    user({ email: 'guest@x.com', external: true }),
    user({ email: 'guestpm@x.com', external: true, role: 'projectManager' }),
    user({ email: 'gone@c.com', role: 'teamManager', active: false })
  ];
  const m = A.securityMembers(users, ['listed@c.com', 'pm@c.com']);
  assert.deepEqual(m.admins, ['listed@c.com', 'pm@c.com', 'adm@c.com']);
  assert.deepEqual(m.approvers, ['guestpm@x.com', 'boss@c.com']);
  assert.deepEqual(m.externals, ['guest@x.com']);
});

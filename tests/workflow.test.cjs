const test = require('node:test');
const assert = require('node:assert/strict');
const { lib, user, project, entry } = require('./helpers.cjs');
const { planReminders } = lib('services/reminderPlan.js');
const { Workflow } = lib('services/workflow.js');
const M = lib('services/mappers.js');
const { defaultConfig } = lib('schema/defaults.js');

test('reminders: past weeks without a submitted sheet, once each, after the reminder day and hour', () => {
  const cfg = { ...defaultConfig(), reminders: { enabled: true, day: 1, hour: 9, currentWeek: false, ccManager: false, lookbackWeeks: 2 } };
  const users = [user({ email: 'a@c.com' }), user({ email: 'b@c.com' }), user({ email: 'c@c.com', active: false })];
  const sheets = [{ userEmail: 'a@c.com', weekStart: '2026-11-16', status: 'Submitted', totalHours: 40 }];
  const now = new Date(2026, 10, 23, 10); // Monday 23 Nov, 10:00
  const plan = planReminders({ cfg, now, users, sheets, projects: [], sent: { 'b@c.com|2026-11-09': 'x' }, remindersOn: true, endDatesOn: false });
  assert.deepEqual(plan.reminders.map((r) => `${r.user.email}|${r.week}`).sort(), ['a@c.com|2026-11-09', 'b@c.com|2026-11-16']);
  const early = planReminders({ cfg, now: new Date(2026, 10, 23, 8), users, sheets, projects: [], sent: {}, remindersOn: true, endDatesOn: false });
  assert.ok(!early.reminders.some((r) => r.week === '2026-11-16'));
});

test('project end notices: the closest threshold once, and on the end date', () => {
  const cfg = { ...defaultConfig(), projectEndNotice: [14, 7, 1] };
  const now = new Date(2026, 10, 23, 10);
  const projects = [project({ id: 1, endDate: '2026-11-28' }), project({ id: 2, endDate: '2026-11-28', endNotices: ['7'] }), project({ id: 3, endDate: '2026-11-20' }), project({ id: 4, endDate: '2027-03-01' })];
  const plan = planReminders({ cfg, now, users: [], sheets: [], projects, sent: {}, remindersOn: false, endDatesOn: true });
  assert.deepEqual(plan.endNotices.map((n) => `${n.project.id}:${n.notice}`), ['1:7', '3:ended']);
});

test('clone a week: lines move to the same weekdays, closed projects dropped, hours optional', () => {
  const wf = new Workflow({ svc: {}, notifier: {}, config: defaultConfig, dir: () => ({}), admins: () => [], has: () => true });
  const src = [entry({ id: 5, projectId: 1, hours: { '2026-11-16': 8, '2026-11-20': 4 }, notes: { '2026-11-16': 'x' } }), entry({ projectId: 2, hours: { '2026-11-17': 1 } })];
  const out = wf.cloneLines(src, '2026-11-16', '2026-11-23', true, [1]);
  assert.equal(out.length, 1);
  assert.equal(out[0].id, 0);
  assert.deepEqual(out[0].hours, { '2026-11-23': 8, '2026-11-27': 4 });
  assert.deepEqual(out[0].notes, {});
  assert.deepEqual(wf.cloneLines(src, '2026-11-16', '2026-11-23', false, [1])[0].hours, {});
});

test('mappers round-trip and date-only columns keep the calendar day', () => {
  const e = entry({ userEmail: 'A@c.com', hours: { '2026-11-16': 7.5, '2026-11-17': 0 }, custom: { t: 'x' } });
  const back = M.toEntry({ Id: 3, ...M.fromEntry(e) });
  assert.deepEqual(back.hours, { '2026-11-16': 7.5 });
  assert.equal(back.total, 7.5);
  assert.equal(M.dateOnly('2026-11-23T18:30:00Z'), '2026-11-24'); // midnight in India, stored as UTC
  assert.equal(M.dateOnly('2026-11-24T08:00:00Z'), '2026-11-24'); // midnight on the US west coast
  assert.deepEqual(M.emails('a@c.com; b@c.com,\nc@c.com'), ['a@c.com', 'b@c.com', 'c@c.com']);
  const p = M.toProject({ Id: 1, Title: 'X', ProjectStatus: 'Weird', MemberRates: '{bad json', IsBillable: null });
  assert.equal(p.status, 'Active');
  assert.deepEqual(p.memberRates, {});
  assert.equal(p.billable, true);
});

const test = require('node:test');
const assert = require('node:assert/strict');
const { lib, user, entry } = require('./helpers.cjs');
const R = lib('services/rules.js');
const { defaultConfig } = lib('schema/defaults.js');
const days = lib('services/dates.js').weekDays('2026-11-16');

test('hours typed in any common form', () => {
  assert.equal(R.parseHours('7.5'), 7.5);
  assert.equal(R.parseHours('7,5'), 7.5);
  assert.equal(R.parseHours('7:30'), 7.5);
  assert.equal(R.parseHours('7h30'), 7.5);
  assert.equal(R.parseHours('45m'), 0.75);
  assert.ok(isNaN(R.parseHours('abc')));
  assert.equal(R.snapHours(7.1, 0.25), 7);
  assert.equal(R.snapHours(30, 0.25), 24);
  assert.equal(R.snapHours(-2, 0.25), 0);
});

test('tracking days: working days, holidays and per-person days', () => {
  const cfg = { ...defaultConfig(), holidays: [{ date: '2026-11-18', name: 'Diwali' }] };
  assert.equal(R.isTrackingDay(cfg, '2026-11-16'), true);
  assert.equal(R.isTrackingDay(cfg, '2026-11-18'), false);
  assert.equal(R.isTrackingDay(cfg, '2026-11-21'), false); // Saturday
  assert.equal(R.isTrackingDay(cfg, '2026-11-21', user({ workDays: [6] })), true);
});

test('organisation limits warn or block', () => {
  const cfg = { ...defaultConfig(), limits: { dailyMin: 0, dailyMax: 10, weeklyMin: 40, weeklyMax: 45, enforcement: 'block' } };
  const e = entry({ projectId: 1, hours: { '2026-11-16': 12, '2026-11-17': 8 }, total: 20 });
  const issues = R.validateWeek({ cfg, days, entries: [e], forSubmit: true });
  assert.ok(issues.some((i) => /daily maximum/.test(i.message) && i.level === 'error'));
  assert.ok(issues.some((i) => /weekly minimum/.test(i.message)));
  assert.equal(R.blocking(issues), true);
  const warn = R.validateWeek({ cfg: { ...cfg, limits: { ...cfg.limits, enforcement: 'warn' } }, days, entries: [e], forSubmit: true });
  assert.equal(warn.filter((i) => /maximum|minimum/.test(i.message)).every((i) => i.level === 'warning'), true);
});

test('rule-based limits: first matching rule, pinned rule wins', () => {
  const contractors = { id: 'c', name: 'Contractors', roles: [], departments: [], emails: [], externalOnly: true, allowNonWorkingDays: false, enforcement: 'block', dailyMin: 0, dailyMax: 8, weeklyMin: 0, weeklyMax: 40 };
  const devs = { ...contractors, id: 'd', name: 'Engineering', externalOnly: false, departments: ['Engineering'], dailyMax: 12 };
  const cfg = { ...defaultConfig(), rulesEnabled: true, rules: [contractors, devs] };
  assert.equal(R.limitsFor(cfg, user({ external: true, department: 'Engineering' })).source, 'Rule: Contractors');
  assert.equal(R.limitsFor(cfg, user({ department: 'engineering' })).dailyMax, 12);
  assert.equal(R.limitsFor(cfg, user({ department: 'Sales' })).source, 'Organisation limits');
  assert.equal(R.limitsFor(cfg, user({ external: true, ruleId: 'd' })).source, 'Rule: Engineering');
  const weekend = R.validateWeek({ cfg, user: user({ external: true }), days, entries: [entry({ hours: { '2026-11-21': 4 }, total: 4 })] });
  assert.ok(weekend.some((i) => /not a working day/.test(i.message) && i.level === 'error'));
});

test('approved leave counts towards minimums unless already on the sheet', () => {
  const cfg = { ...defaultConfig(), limits: { dailyMin: 8, dailyMax: 0, weeklyMin: 40, weeklyMax: 0, enforcement: 'block' } };
  const work = entry({ projectId: 1, hours: { '2026-11-16': 8, '2026-11-17': 8, '2026-11-18': 8, '2026-11-19': 8 }, total: 32 });
  const leave = [{ email: 'p@contoso.com', start: '2026-11-20', end: '2026-11-20', type: 'Annual' }];
  assert.equal(R.blocking(R.validateWeek({ cfg, days, entries: [work], timeOff: leave, forSubmit: true })), false);
  assert.equal(R.blocking(R.validateWeek({ cfg, days, entries: [work], forSubmit: true })), true);
});

test('submit checks: project, activity, notes, required custom columns', () => {
  const cfg = { ...defaultConfig(), requireNotes: true };
  const e = entry({ projectId: 0, hours: { '2026-11-16': 2 }, total: 2 });
  const msgs = R.validateWeek({ cfg, days, entries: [e], forSubmit: true, requireActivity: true, requiredColumns: [{ key: 'ticket', label: 'Ticket #' }] }).map((i) => i.message);
  assert.ok(msgs.some((m) => /needs a project/.test(m)));
  assert.ok(msgs.some((m) => /needs an activity/.test(m)));
  assert.ok(msgs.some((m) => /add a note/.test(m)));
  assert.ok(msgs.some((m) => /Ticket #/.test(m)));
  // a time-off line needs no project
  const off = entry({ projectId: 0, activityId: 9, hours: { '2026-11-16': 8 }, total: 8 });
  assert.ok(!R.validateWeek({ cfg: defaultConfig(), days, entries: [off], forSubmit: true, timeOffActivityId: 9 }).some((i) => /project/.test(i.message)));
});

test('leave spanning a weekend counts only on working days', () => {
  const cfg = { ...defaultConfig(), limits: { dailyMin: 0, dailyMax: 0, weeklyMin: 0, weeklyMax: 45, enforcement: 'block' } };
  const work = entry({ projectId: 1, hours: { '2026-11-16': 8, '2026-11-17': 8 }, total: 16 });
  const leave = [{ email: 'p@contoso.com', start: '2026-11-18', end: '2026-11-23', type: 'Annual' }];
  // Wed–Fri leave = 24 h; Sat and Sun must not add another 16 h (which would break the 45 h maximum).
  assert.equal(R.blocking(R.validateWeek({ cfg, days, entries: [work], timeOff: leave, forSubmit: true })), false);
});

test('a public holiday on a working day lowers the weekly minimum by a standard day', () => {
  const cfg = { ...defaultConfig(), holidays: [{ date: '2026-11-18', name: 'Diwali' }, { date: '2026-11-21', name: 'Saturday event' }], limits: { dailyMin: 0, dailyMax: 0, weeklyMin: 40, weeklyMax: 0, enforcement: 'block' } };
  assert.equal(R.weeklyTarget(cfg, days), 32); // Wednesday counts, the Saturday holiday does not
  const work = entry({ projectId: 1, hours: { '2026-11-16': 8, '2026-11-17': 8, '2026-11-19': 8, '2026-11-20': 8 }, total: 32 });
  assert.equal(R.blocking(R.validateWeek({ cfg, days, entries: [work], forSubmit: true })), false);
  const short = entry({ projectId: 1, hours: { '2026-11-16': 8, '2026-11-17': 8, '2026-11-19': 8 }, total: 24 });
  assert.ok(R.validateWeek({ cfg, days, entries: [short], forSubmit: true }).some((i) => /weekly minimum of 32 h \(40 h less public holidays\)/.test(i.message)));
  assert.equal(R.weeklyTarget({ ...cfg, limits: { ...cfg.limits, weeklyMin: 0 } }, days), 32); // no minimum: 4 tracking days × 8
});

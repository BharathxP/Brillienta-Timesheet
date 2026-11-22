const test = require('node:test');
const assert = require('node:assert/strict');
const { lib, user, project, entry, sheet } = require('./helpers.cjs');
const R = lib('services/reports.js');

const u1 = user({ email: 'a@c.com', name: 'Asha', department: 'Eng' });
const u2 = user({ email: 'b@c.com', name: 'Bilal', department: 'Ops' });
const p = project({ id: 1, title: 'Web', client: 'Contoso' });
const q = project({ id: 2, title: 'Internal' });
const lk = { users: [u1, u2], projects: [p, q], tasks: [], activities: [], programs: [] };
const s1 = sheet({ id: 10, status: 'Approved' });
const s2 = sheet({ id: 11, status: 'Draft' });
const entries = [
  entry({ sheetId: 10, userEmail: 'a@c.com', projectId: 1, billable: true, hours: { '2026-11-16': 8, '2026-11-17': 4 }, costRate: 100, billRate: 300, custom: { ticket: 'T1' } }),
  entry({ sheetId: 11, userEmail: 'b@c.com', projectId: 2, billable: false, hours: { '2026-11-16': 2, '2026-12-01': 5 } })
];

test('facts: one per person, line and day, filtered by dates, billable and status', () => {
  const all = R.toFacts(entries, [s1, s2], lk, { from: '2026-11-01', to: '2026-11-30' }, 1);
  assert.equal(all.length, 3);
  assert.equal(R.toFacts(entries, [s1, s2], lk, { billable: 'billable' }, 1).length, 2);
  assert.equal(R.toFacts(entries, [s1, s2], lk, { statuses: ['Approved'] }, 1).length, 2);
  assert.equal(R.toFacts(entries, [s1, s2], lk, { userEmails: ['B@c.com'] }, 1).length, 2);
});

test('group and total: hours, billable split, cost, revenue, margin', () => {
  const facts = R.toFacts(entries, [s1, s2], lk, {}, 1);
  const t = R.totals(facts);
  assert.equal(t.hours, 19);
  assert.equal(t.billableHours, 12);
  assert.equal(t.nonBillableHours, 7);
  assert.equal(t.cost, 1200);
  assert.equal(t.revenue, 3600);
  assert.equal(t.margin, 2400);
  const byProject = R.groupFacts(facts, ['project'], lk);
  assert.deepEqual(byProject.map((r) => [r.labels[0], r.values.hours]), [['Internal', 7], ['Web', 12]]);
  const byMonth = R.groupFacts(facts, ['month', 'user'], lk);
  assert.deepEqual(byMonth.map((r) => r.labels.join('/')), ['Nov 2026/Asha', 'Nov 2026/Bilal', 'Dec 2026/Bilal']);
  assert.deepEqual(R.groupFacts(facts, ['custom:ticket'], lk).map((r) => r.labels[0]).sort(), ['(empty)', 'T1']);
});

test('drafts without captured rates use the current rates', () => {
  const facts = R.toFacts(entries, [s1, s2], { ...lk, rateFor: () => ({ cost: 50, bill: 0 }) }, {}, 1);
  assert.equal(R.totals(facts).cost, 1200 + 7 * 50);
});

test('utilisation', () => {
  assert.equal(R.utilisation(30, 40), 75);
  assert.equal(R.utilisation(5, 0), 0);
});

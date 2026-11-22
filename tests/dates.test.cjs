const test = require('node:test');
const assert = require('node:assert/strict');
const { lib } = require('./helpers.cjs');
const D = lib('services/dates.js');

test('week start honours the configured first day of the week', () => {
  assert.equal(D.weekStartOf('2026-11-22', 1), '2026-11-16'); // Sunday → Monday week
  assert.equal(D.weekStartOf('2026-11-22', 0), '2026-11-22'); // Sunday week
  assert.equal(D.weekStartOf('2026-11-20', 6), '2026-11-14'); // Saturday week
  assert.deepEqual(D.weekDays('2026-12-28').slice(-2), ['2027-01-02', '2027-01-03']);
});

test('periods and weeks in a range', () => {
  assert.deepEqual(D.periodRange('lastMonth', '2026-01-15', 1), { from: '2025-12-01', to: '2025-12-31' });
  assert.deepEqual(D.periodRange('thisQuarter', '2026-11-22', 1), { from: '2026-10-01', to: '2026-12-31' });
  assert.deepEqual(D.weeksIn('2026-11-01', '2026-11-20', 1), ['2026-10-26', '2026-11-02', '2026-11-09', '2026-11-16']);
  assert.equal(D.daysBetween('2026-03-28', '2026-03-30'), 2); // across a DST change
});

test('labels', () => {
  assert.equal(D.formatWeek('2026-11-16'), '16 – 22 Nov 2026');
  assert.equal(D.formatWeek('2026-12-28'), '28 Dec 2026 – 3 Jan 2027');
  assert.equal(D.isoWeek('2026-11-16'), 47);
});

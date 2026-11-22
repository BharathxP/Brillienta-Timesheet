const test = require('node:test');
const assert = require('node:assert/strict');
const { lib } = require('./helpers.cjs');
const B = lib('services/bulk.js');

test('rows pasted from Excel, with a header in another order', () => {
  const rows = B.parseBulk('projects', 'Code\tProject\tBillable\tManagers\nWEB\tWebsite\tyes\ta@c.com; b@c.com\nAPP\t\tmaybe\tnot-an-email');
  assert.equal(rows.length, 2);
  assert.deepEqual(rows[0].values.managers, ['a@c.com', 'b@c.com']);
  assert.equal(rows[0].values.code, 'WEB');
  assert.equal(rows[0].values.billable, true);
  assert.deepEqual(rows[0].errors, []);
  assert.ok(rows[1].errors.some((e) => /required/.test(e)));
  assert.ok(rows[1].errors.some((e) => /yes or no/.test(e)));
  assert.ok(rows[1].errors.some((e) => /Not an email/.test(e)));
});

test('CSV without a header follows the template order; dates and numbers are checked', () => {
  const rows = B.parseBulk('tasks', 'Web,"Design, phase 1",,yes,40,2026-12-01,no\nWeb,Build,,,x,01/12/2026,');
  assert.equal(rows[0].values.title, 'Design, phase 1');
  assert.equal(rows[0].values.planned, 40);
  assert.equal(rows[0].values.selfAssign, false);
  assert.ok(rows[1].errors.some((e) => /number/.test(e)));
  assert.ok(rows[1].errors.some((e) => /yyyy-mm-dd/.test(e)));
  assert.equal(B.parseBulk('activities', 'Activity\tBillable\nTravel\tno')[0].values.billable, false);
});

const test = require('node:test');
const assert = require('node:assert/strict');
const nodeCrypto = require('crypto');
const { lib } = require('./helpers.cjs');
const L = lib('services/licensing/license.js');
const P = lib('services/licensing/plans.js');
const { defaultConfig } = lib('schema/defaults.js');

const TENANT = '11111111-2222-4333-8444-555555555555';
const { privateKey, publicKey } = nodeCrypto.generateKeyPairSync('ec', { namedCurve: 'P-256' });
const PUB = publicKey.export({ type: 'spki', format: 'der' }).toString('base64');
const b64url = (b) => Buffer.from(b).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const subtle = globalThis.crypto.subtle;
const DAY = 86400000;
function issue(payload, prefix = 'BT1') {
  const signed = prefix + '.' + b64url(Buffer.from(JSON.stringify({ v: 1, product: 'timesheet', id: 'L-1', customer: 'Contoso', users: 0, issued: '2026-10-01', ...payload })));
  return signed + '.' + b64url(nodeCrypto.sign('sha256', Buffer.from(signed, 'ascii'), { key: privateKey, dsaEncoding: 'ieee-p1363' }));
}

test('a genuine key for this tenant is accepted; tampering, other tenants and other products are not', async () => {
  const key = issue({ plan: 'pro', tenant: TENANT, users: 45 });
  assert.equal((await L.checkLicenseKey(key, TENANT.toUpperCase(), PUB, subtle)).users, 45);
  await assert.rejects(L.checkLicenseKey(key, '99999999-2222-4333-8444-555555555555', PUB, subtle), /another Microsoft 365/);
  const parts = key.split('.');
  const forged = parts[0] + '.' + b64url(Buffer.from(JSON.stringify({ ...JSON.parse(Buffer.from(parts[1], 'base64')), plan: 'enterprise' }))) + '.' + parts[2];
  await assert.rejects(L.checkLicenseKey(forged, TENANT, PUB, subtle), /not genuine/);
  await assert.rejects(L.checkLicenseKey(issue({ plan: 'pro', tenant: TENANT }, 'ED1'), TENANT, PUB, subtle), /Employee Directory/);
});

test('trial, licence, grace and end', () => {
  const start = new Date('2026-11-01T00:00:00Z');
  assert.equal(L.licenseStatus(undefined, start, new Date(start.getTime() + 3 * DAY)).plan, 'enterprise');
  assert.equal(L.licenseStatus(undefined, start, new Date(start.getTime() + 15 * DAY)).state, 'trialEnded');
  const p = { v: 1, product: 'timesheet', id: 'x', plan: 'standard', tenant: TENANT, customer: 'C', users: 0, issued: '2026-01-01', expires: '2026-12-31' };
  assert.equal(L.licenseStatus(p, start, new Date('2026-12-01')).users, 10);
  assert.equal(L.licenseStatus(p, start, new Date('2027-01-05')).state, 'grace');
  assert.equal(L.licenseStatus(p, start, new Date('2027-02-01')).plan, undefined);
});

test('plans: features, limits, add-ons', () => {
  assert.equal(P.planHas('standard', 'cloneWeek'), true);
  assert.equal(P.planHas('standard', 'autoApproval'), false);
  assert.equal(P.planHas('pro', 'autoApproval'), true);
  assert.equal(P.planHas('pro', 'planner'), false);
  assert.equal(P.planHas('pro', 'externalUsers'), false);
  assert.equal(P.planHas('pro', 'externalUsers', ['externalUsers']), true);
  assert.equal(P.planHas('standard', 'externalUsers', ['externalUsers']), false);
  assert.equal(P.userLimit('pro'), 30);
  assert.equal(P.userLimit('pro', 42), 42);
  assert.equal(P.ticketLimit('standard'), 4);
  assert.equal(P.ticketLimit('pro'), 0);
});

test('a plan switches off what it does not include, and saving keeps locked settings', () => {
  const saved = { ...defaultConfig(), weekStart: 0, rulesEnabled: true, rules: [{ id: 'r' }], autoApproval: { enabled: true, roles: ['admin'], requireWithinRules: true, maxHours: 0 } };
  const std = P.applyPlan(saved, 'standard');
  assert.equal(std.weekStart, 1);
  assert.equal(std.rulesEnabled, false);
  assert.equal(std.autoApproval.enabled, false);
  assert.equal(P.applyPlan(saved, 'enterprise').weekStart, 0);
  const next = P.restoreLocked({ ...std, orgName: 'New' }, saved, 'standard');
  assert.equal(next.orgName, 'New');
  assert.equal(next.weekStart, 0);
  assert.equal(next.rules.length, 1);
});

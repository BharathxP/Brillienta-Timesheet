#!/usr/bin/env node
// Brillienta Timesheet licence tool (vendor only; never shipped in the .sppkg).
//
// Timesheet keys are signed with the same key pair as Brillienta Employee Directory (its public key is in
// src/webparts/brillientaTimesheet/services/licensing/publicKey.ts), so the private key you already keep for the
// directory issues them. Keys carry "product":"timesheet" and the BT1 prefix, so a directory key never unlocks Timesheet.
//
//   node tools/license/license.cjs issue --plan standard|pro|enterprise --tenant <tenant id> --customer "Contoso Ltd"
//        [--users 45] [--expires 2027-12-31] [--addons externalUsers] [--key <private-key.pem>] [--id <licence id>]
//        [--am-name "Kavya Reddy" --am-email kavya@brillienta.com --am-phone "+91 …"]
//       Prints a licence key for one Microsoft 365 tenant (customers find the tenant id in Admin › Licence & plan).
//       --users 0 (default) = the plan's standard number (Standard 10, Pro 30, Enterprise 30); extra users are sold
//       at ₹200 (Pro) / ₹300 (Enterprise) per user: issue the total. Leave out --expires for a perpetual licence.
//       --am-* (Enterprise): the dedicated account manager shown on the customer's Support page.
//
//   node tools/license/license.cjs verify <licence key> [--tenant <tenant id>]
//       Checks a key against the public key in the app and prints what it contains.
//
// Keys: ECDSA P-256 / SHA-256. Format: BT1.<base64url(payload JSON)>.<base64url(signature, IEEE P1363)>

const crypto = require('crypto');
const fs = require('fs');
const os = require('os');
const path = require('path');

const PUBLIC_KEY_FILE = path.join(__dirname, '..', '..', 'src', 'webparts', 'brillientaTimesheet', 'services', 'licensing', 'publicKey.ts');
const DEFAULT_KEY = path.join(os.homedir(), '.employee-directory-licensing', 'private-key.pem');
const PLANS = ['standard', 'pro', 'enterprise'];
const ADDONS = ['externalUsers'];
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const b64url = (buf) => Buffer.from(buf).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const fromB64url = (s) => Buffer.from(s.replace(/-/g, '+').replace(/_/g, '/'), 'base64');

function args(argv) {
  const out = { _: [] };
  for (let i = 0; i < argv.length; i++) {
    if (argv[i].startsWith('--')) out[argv[i].slice(2)] = argv[i + 1] && !argv[i + 1].startsWith('--') ? argv[++i] : 'true';
    else out._.push(argv[i]);
  }
  return out;
}

function fail(msg) {
  console.error('Error: ' + msg);
  process.exit(1);
}

function appPublicKey() {
  const src = fs.readFileSync(PUBLIC_KEY_FILE, 'utf8');
  const m = /LICENSE_PUBLIC_KEY\s*=\s*'([^']+)'/.exec(src);
  if (!m) fail('No public key found in ' + PUBLIC_KEY_FILE);
  return crypto.createPublicKey({ key: Buffer.from(m[1], 'base64'), format: 'der', type: 'spki' });
}

function issue(a) {
  if (PLANS.indexOf(a.plan) < 0) fail('--plan must be one of ' + PLANS.join(', '));
  if (!GUID.test(a.tenant || '')) fail('--tenant must be the customer’s Microsoft 365 tenant id (a GUID)');
  if (!a.customer) fail('--customer is required');
  if (a.expires && !/^\d{4}-\d{2}-\d{2}$/.test(a.expires)) fail('--expires must be yyyy-mm-dd');
  const addons = (a.addons || '').split(',').map((s) => s.trim()).filter((s) => !!s);
  addons.forEach((x) => ADDONS.indexOf(x) < 0 && fail('Unknown add-on ' + x + ' (known: ' + ADDONS.join(', ') + ')'));
  const keyFile = a.key || DEFAULT_KEY;
  if (!fs.existsSync(keyFile)) fail('Private key not found: ' + keyFile + ' (pass --key <private-key.pem>)');
  const privateKey = crypto.createPrivateKey(fs.readFileSync(keyFile));
  const derived = crypto.createPublicKey(privateKey).export({ type: 'spki', format: 'der' }).toString('base64');
  if (derived !== appPublicKey().export({ type: 'spki', format: 'der' }).toString('base64')) fail('This private key does not match the public key built into the app; keys it signs would be rejected.');
  const payload = {
    v: 1,
    product: 'timesheet',
    id: a.id || 'BT-' + Date.now().toString(36).toUpperCase(),
    plan: a.plan,
    tenant: a.tenant.toLowerCase(),
    customer: a.customer,
    users: Number(a.users || 0),
    issued: new Date().toISOString().substring(0, 10)
  };
  if (a.expires) payload.expires = a.expires;
  if (addons.length) payload.addons = addons;
  if (a['am-name']) payload.accountManager = { name: a['am-name'], email: a['am-email'] || '', phone: a['am-phone'] || '' };
  const signed = 'BT1.' + b64url(Buffer.from(JSON.stringify(payload)));
  const sig = crypto.sign('sha256', Buffer.from(signed, 'ascii'), { key: privateKey, dsaEncoding: 'ieee-p1363' });
  console.log(signed + '.' + b64url(sig));
  console.error(`\nLicence ${payload.id}: ${payload.plan} for ${payload.customer} (${payload.tenant})${payload.users ? ', ' + payload.users + ' users' : ''}${payload.expires ? ', until ' + payload.expires : ', perpetual'}`);
}

function verify(key, a) {
  const parts = (key || '').replace(/\s+/g, '').split('.');
  if (parts.length !== 3 || parts[0] !== 'BT1') fail('Not a Brillienta Timesheet key (BT1.…)');
  const payload = JSON.parse(fromB64url(parts[1]).toString('utf8'));
  const ok = crypto.verify('sha256', Buffer.from(parts[0] + '.' + parts[1], 'ascii'), { key: appPublicKey(), dsaEncoding: 'ieee-p1363' }, fromB64url(parts[2]));
  console.log(JSON.stringify(payload, null, 2));
  console.log('Signature: ' + (ok ? 'valid' : 'INVALID'));
  if (a.tenant) console.log('Tenant: ' + (payload.tenant.toLowerCase() === a.tenant.toLowerCase() ? 'matches' : 'DOES NOT MATCH'));
  if (!ok) process.exit(2);
}

const a = args(process.argv.slice(2));
const cmd = a._[0];
if (cmd === 'issue') issue(a);
else if (cmd === 'verify') verify(a._[1], a);
else {
  console.log(fs.readFileSync(__filename, 'utf8').split('\n').filter((l) => l.startsWith('//')).map((l) => l.slice(3)).join('\n'));
}

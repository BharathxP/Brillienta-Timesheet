const test = require('node:test');
const assert = require('node:assert/strict');
const { lib } = require('./helpers.cjs');
const K = lib('services/packages.js');
const { zipFiles } = lib('services/zip.js');

test('Copilot agent is grounded on the timesheet site and names its lists', () => {
  const o = { siteUrl: 'https://contoso.sharepoint.com/sites/ts', orgName: 'Contoso', appName: 'Brillienta Timesheet', prefix: 'BrillTS_', appId: '00000000-0000-4000-8000-000000000000', version: '1.0.0' };
  const a = K.declarativeAgent(o);
  assert.deepEqual(a.capabilities[0].items_by_url, [{ url: o.siteUrl }]);
  assert.match(a.instructions, /BrillTS_TimeEntries/);
  assert.equal(K.copilotManifest(o).copilotAgents.declarativeAgents[0].file, 'declarativeAgent.json');
  assert.ok(a.instructions.length < 8000);
});

test('Power BI data source and API endpoints', () => {
  assert.equal(JSON.parse(K.pbids('https://x/sites/ts')).connections[0].details.protocol, 'sharepoint-list');
  const e = K.endpoints('https://x/sites/ts/', 'BrillTS_');
  assert.equal(e.length, 13);
  assert.match(e.filter((x) => x.key === 'entries')[0].rest, /^https:\/\/x\/sites\/ts\/_api\/web\/lists\/getbytitle\('BrillTS_TimeEntries'\)\/items$/);
});

test('zip takes text and binary files', () => {
  const z = zipFiles([{ path: 'a.txt', content: 'hi' }, { path: 'b.bin', content: new Uint8Array([1, 2, 3]) }]);
  assert.equal(z[0], 0x50);
  assert.equal(z[1], 0x4b);
});

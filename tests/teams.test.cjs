// Publication: what the SharePoint Store / AppSource listing, Sync to Teams and Teams admin center need from the
// package, the in-app Teams and Copilot downloads, and how Teams finds the timesheet site.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { lib } = require('./helpers.cjs');
const H = lib('services/host.js');
const K = lib('services/packages.js');
const { PRODUCT, WEB_PART_ID } = lib('services/product.js');

const root = path.resolve(__dirname, '..');
const manifest = JSON.parse(fs.readFileSync(path.join(root, 'src/webparts/brillientaTimesheet/BrillientaTimesheetWebPart.manifest.json'), 'utf8'));
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'config/package-solution.json'), 'utf8')).solution;

function pngSize(file) {
  const b = fs.readFileSync(file);
  assert.equal(b.toString('ascii', 1, 4), 'PNG', `${file} is a PNG`);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20), colorType: b[25] };
}

test('store listing: the package says what services/product.ts says', () => {
  assert.equal(pkg.name, PRODUCT.name, 'solution name = marketplace name');
  assert.deepEqual(pkg.developer, PRODUCT.developer);
  assert.equal(pkg.metadata.shortDescription.default, PRODUCT.shortDescription);
  assert.equal(pkg.metadata.longDescription.default, PRODUCT.longDescription);
  assert.equal(manifest.id, WEB_PART_ID);
});

test('store listing: descriptions within the limits, ending with Get started, Help and Contact us', () => {
  assert.ok(PRODUCT.shortDescription.length <= 80, 'short description ≤ 80');
  assert.ok(PRODUCT.longDescription.length <= 4000, 'long description ≤ 4000');
  const d = PRODUCT.longDescription;
  for (const s of ['Get started', 'Help', 'Contact us', PRODUCT.developer.websiteUrl, PRODUCT.helpUrl, PRODUCT.contactUrl, PRODUCT.supportEmail, 'Sync to Teams']) assert.ok(d.includes(s), s);
  // Every Graph permission the package requests is explained in the description.
  for (const p of pkg.webApiPermissionRequests) assert.ok(d.includes(p.scope), `${p.scope} explained`);
});

test('store listing: app icon and five 1366 × 768 screenshots in the package', () => {
  const icon = pngSize(path.join(root, 'sharepoint', pkg.iconPath));
  assert.deepEqual([icon.w, icon.h], [96, 96], 'app icon 96 × 96');
  assert.equal(pkg.metadata.screenshotPaths.length, 5);
  for (const s of pkg.metadata.screenshotPaths) {
    const size = pngSize(path.join(root, 'sharepoint', s));
    assert.deepEqual([size.w, size.h], [1366, 768], `${s} is 1366 × 768`);
  }
  const allowed = ['Accounting + Finance', 'Collaboration', 'Content management', 'CRM', 'Data + analytics', 'File managers', 'IT/admin', 'Legal + HR', 'News + weather', 'Productivity', 'Project management', 'Reference', 'Sales + marketing', 'Site Design', 'Social', 'Workflow & Process Management'];
  for (const c of pkg.metadata.categories) assert.ok(allowed.indexOf(c) >= 0, `category ${c}`);
});

test('store listing: publisher details filled in (Teams rejects an app without them)', () => {
  const d = pkg.developer;
  assert.ok(d.name && d.name.length <= 32);
  for (const k of ['websiteUrl', 'privacyUrl', 'termsOfUseUrl']) assert.match(d[k], /^https:\/\/\S+$/, `${k} is an https URL`);
  assert.ok(d.mpnId, 'partner id');
});

test('web part picker: the app logo, the marketplace name and a short description', () => {
  const icon = 'data:image/png;base64,' + fs.readFileSync(path.join(root, 'sharepoint', pkg.iconPath)).toString('base64');
  for (const e of manifest.preconfiguredEntries) {
    assert.equal(e.iconImageUrl, icon, 'the app icon, not a font icon');
    assert.equal(e.officeFabricIconFontName, undefined);
    assert.equal(e.title.default, PRODUCT.name);
    assert.ok(e.description.default.length <= 80);
  }
  assert.ok(manifest.supportedHosts.includes('TeamsTab') && manifest.supportedHosts.includes('TeamsPersonalApp'));
  assert.equal(pkg.skipFeatureDeployment, true, 'tenant-wide (needed by Sync to Teams)');
});

test('versions agree: package.json, the solution and its feature', () => {
  const app = require('../package.json').version;
  assert.equal(pkg.version, app + '.0');
  for (const f of pkg.features) assert.equal(f.version, app + '.0');
});

test('Teams app: teams/manifest.json is current, valid against the official schema, with tier1 channel support', () => {
  const T = require('../tools/teams/app-manifest.cjs');
  const text = fs.readFileSync(path.join(root, 'teams', 'manifest.json'), 'utf8');
  assert.equal(text, T.manifestText(), 'regenerate with: node tools/teams/make-app-package.cjs');
  const shipped = JSON.parse(text);
  const icons = T.iconNames(manifest.id);
  assert.deepEqual(shipped.icons, icons);
  const color = fs.readFileSync(path.join(root, 'teams', icons.color));
  const outline = fs.readFileSync(path.join(root, 'teams', icons.outline));
  assert.deepEqual(T.checkPackage(shipped, color, outline), []);
  assert.equal(shipped.name.short, pkg.name);
  assert.equal(shipped.version, require('../package.json').version);
  assert.equal(shipped.id, manifest.id, 'app id = web part id, so Sync to Teams updates the same app');
  assert.equal(shipped.supportsChannelFeatures, 'tier1');
  for (const t of shipped.staticTabs.concat(shipped.configurableTabs)) assert.ok((t.contentUrl || t.configurationUrl).includes(`componentId=${manifest.id}`));
  // The in-app download (Admin › Integrations › Teams) is the very same manifest.
  assert.deepEqual(K.teamsManifest(shipped.version), shipped);
  const zip = T.buildPackage();
  for (const name of ['manifest.json', icons.color, icons.outline]) assert.ok(zip.includes(Buffer.from(name)), `${name} in the upload zip`);
});

test('Teams app: the checker rejects broken manifests and icons', () => {
  const T = require('../tools/teams/app-manifest.cjs');
  const s = T.sources();
  const bad = T.buildManifest(s);
  bad.name.short = 'x'.repeat(31);
  bad.extra = true;
  delete bad.webApplicationInfo;
  delete bad.supportsChannelFeatures;
  const problems = T.checkPackage(bad, s.outline, s.color).join('\n');
  for (const p of ['name/short', '(extra)', 'webApplicationInfo', 'supportsChannelFeatures', 'colour icon is 32', 'outline icon is 192']) assert.ok(problems.includes(p), p);
});

test('Copilot agent package: valid against the official Teams and declarative agent schemas; stable id per site', () => {
  const T = require('../tools/teams/app-manifest.cjs');
  const o = { siteUrl: 'https://contoso.sharepoint.com/sites/timesheet', orgName: 'Contoso', appName: 'Brillienta Timesheet', prefix: 'BrillTS_', version: '1.0.0' };
  const m = K.copilotManifest(o);
  assert.deepEqual(T.checkManifestOnly(m), []);
  assert.deepEqual(T.checkDeclarativeAgent(K.declarativeAgent(o)), []);
  assert.equal(K.copilotManifest(o).id, m.id, 'downloading again gives the same app id (update, not a second agent)');
  assert.notEqual(K.copilotManifest({ ...o, siteUrl: o.siteUrl + '2' }).id, m.id);
  assert.match(m.id, /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/);
  assert.deepEqual(m.developer, PRODUCT.developer);
});

test('Teams finds the timesheet with SharePoint search: Settings list rows → sites, one pick', () => {
  const o = 'https://contoso.sharepoint.com';
  const row = (title, p, web) => ({ Cells: [{ Key: 'Title', Value: title }, { Key: 'Path', Value: p }, { Key: 'SPWebUrl', Value: web }] });
  const rows = [
    row('BrillTS_Settings', 'https://contoso.sharepoint.com/sites/ts/Lists/BrillTS_Settings/AllItems.aspx', 'https://contoso.sharepoint.com/sites/ts'),
    row('Something', 'https://contoso.sharepoint.com/sites/ts2/Lists/BrillTS_Settings', null),
    row('BrillTS_SettingsOld', 'https://contoso.sharepoint.com/sites/x/Lists/BrillTS_SettingsOld', 'https://contoso.sharepoint.com/sites/x'),
    row('BrillTS_Settings', 'https://contoso.sharepoint.com/sites/ts/Lists/BrillTS_Settings/1_.000', 'https://contoso.sharepoint.com/sites/ts'),
    row('BrillTS_Settings', 'https://fabrikam.sharepoint.com/sites/y/Lists/BrillTS_Settings', 'https://fabrikam.sharepoint.com/sites/y')
  ];
  assert.deepEqual(H.sitesFromSearch(rows, 'BrillTS_Settings', o), ['https://contoso.sharepoint.com/sites/ts', 'https://contoso.sharepoint.com/sites/ts2']);
  assert.equal(H.settingsSearchQuery("BrillTS_Settings' OR x"), 'BrillTS_SettingsORx contentclass:STS_List*');
  const a = 'https://contoso.sharepoint.com/sites/a';
  const b = 'https://contoso.sharepoint.com/sites/b';
  assert.equal(H.pickSite([], ''), undefined);
  assert.deepEqual(H.pickSite([{ url: a, marked: false }], ''), { url: a });
  assert.deepEqual(H.pickSite([{ url: a, marked: false }, { url: b, marked: true }], ''), { url: b }, 'the one marked "Use this site in Teams"');
  assert.deepEqual(H.pickSite([{ url: a, marked: false }, { url: b, marked: true }], a + '/'), { url: a }, "this browser's earlier choice wins");
  assert.deepEqual(H.pickSite([{ url: a, marked: false }, { url: b, marked: false }], ''), { choose: [a, b] });
});

test('lists out of search: operational lists hidden from search; Settings and the Copilot lists stay searchable', () => {
  const { LISTS } = lib('schema/schema.js');
  const crawl = (k) => !LISTS.filter((l) => l.key === k)[0].noCrawl;
  for (const k of ['settings', 'projects', 'tasks', 'timesheets', 'entries', 'activities', 'users']) assert.ok(crawl(k), `${k} searchable`);
  for (const k of ['audit', 'emailQueue', 'notifications', 'tickets', 'reports']) assert.ok(!crawl(k), `${k} out of search`);
});

test('no Employee Directory leftovers in what customers see', () => {
  const flow = lib('services/flowPackage.js');
  const files = flow.flowPackageFiles({ siteUrl: 'https://t/sites/ts', listId: 'x', listTitle: 'BrillTS_EmailQueue', displayName: 'Flow', sharedMailbox: '', allowedDomains: [], useReplyTo: false });
  const text = files.map((f) => f.content).join('\n');
  assert.ok(!/employee directory|directory's emails/i.test(text), 'flow package mentions the directory');
  assert.ok(!/directory/i.test(PRODUCT.longDescription));
});

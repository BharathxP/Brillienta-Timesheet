// Captures the marketplace screenshots (sharepoint/images/screenshot-1..5.png, 1366 × 768, as the SharePoint Store and
// AppSource require) and the product guide screenshots (docs/client/shots) from the preview, with an Enterprise
// licence so no trial banner shows.
// Usage: npm run preview, then node tools/docs/capture.cjs [store|guide] [name filter]
const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

const root = path.resolve(__dirname, '../..');
const fonts = path.join(root, 'node_modules/@fluentui/font-icons-mdl2/fonts');
const which = process.argv[2] || 'all';
const only = process.argv[3] ? new RegExp(process.argv[3]) : undefined;
const Q = '?role=admin&plan=enterprise&demo=1';

// name, route, query, wait-for text (optional), action (optional)
const STORE = [
  ['screenshot-1', 'home', Q],
  ['screenshot-2', 'timesheet', Q],
  ['screenshot-3', 'approvals', Q],
  ['screenshot-4', 'projects', Q],
  ['screenshot-5', 'reports', Q]
].map((s) => ({ name: s[0], route: s[1], query: s[2], dir: path.join(root, 'sharepoint/images'), w: 1366, h: 768 }));
const GUIDE = [
  ['home', 'home', Q],
  ['week', 'timesheet', Q],
  ['history', 'history', Q],
  ['tasks', 'tasks', Q],
  ['approvals', 'approvals', Q],
  ['projects', 'projects', Q],
  ['programs', 'programs', Q],
  ['activities', 'activities', Q],
  ['people', 'people', Q],
  ['reports', 'reports', Q],
  ['dashboards', 'dashboards', Q],
  ['admin', 'admin', Q],
  ['admin-integrations', 'admin/integrations', Q],
  ['licence', 'admin/licence', Q],
  ['support', 'support', Q]
].map((s) => ({ name: s[0], route: s[1], query: s[2], dir: path.join(root, 'docs/client/shots'), w: 1366, h: 860 }));
const shots = (which === 'store' ? STORE : which === 'guide' ? GUIDE : STORE.concat(GUIDE)).filter((s) => !only || only.test(s.name));

(async () => {
  const browser = await chromium.launch();
  const base = 'file://' + path.join(root, 'preview-dist').split(path.sep).join('/');
  for (const s of shots) {
    fs.mkdirSync(s.dir, { recursive: true });
    const page = await browser.newPage({ viewport: { width: s.w, height: s.h }, deviceScaleFactor: 1 });
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    // Fluent UI's icon font normally comes from Microsoft's CDN: serve the installed copy, so icons show offline too.
    await page.route(/\.woff2?(\?.*)?$/, (route) => {
      const file = path.join(fonts, path.basename(new URL(route.request().url()).pathname));
      return fs.existsSync(file) ? route.fulfill({ path: file, contentType: 'font/woff' }) : route.continue();
    });
    await page.goto(`${base}/index.html${s.query}#/${s.route}`);
    await page.waitForTimeout(3500);
    // Toasts from loading (e.g. "Saved") would date the picture.
    await page.addStyleTag({ content: '.ts-toasts{display:none!important}' });
    await page.screenshot({ path: path.join(s.dir, s.name + '.png') });
    await page.close();
    if (errors.length) throw new Error(`${s.name}: ${errors.join('; ')}`);
    console.log(path.relative(root, path.join(s.dir, s.name + '.png')));
  }
  await browser.close();
})().catch((e) => {
  console.error(e.message);
  process.exit(1);
});

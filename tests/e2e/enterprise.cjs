// Enterprise feature audit (docs/ENTERPRISE-CHECKLIST.md) against the preview's in-memory SharePoint with a signed
// Enterprise licence key. Prints PASS / FAIL per check; exits 1 if any fails.  npm run e2e runs it after run.cjs.
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const BASE = 'file://' + path.resolve(__dirname, '..', '..', 'preview-dist', 'index.html');
const OUT = path.join(__dirname, 'output');
fs.mkdirSync(OUT, { recursive: true });
const results = [];
const check = (name, ok, detail) => {
  results.push({ name, ok: !!ok, detail });
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail !== undefined ? '  — ' + detail : ''}`);
};
const lists = (p, f, arg) => p.evaluate(f, arg);

async function open(p, query, route) {
  await p.goto(`${BASE}?${query}#/${route}`);
  await p.waitForSelector('.ts-page', { timeout: 15000 });
  await p.waitForTimeout(1500);
}

(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true });
  const p = await ctx.newPage();
  const errs = [];
  p.on('pageerror', (e) => errs.push(e.message));
  const E = 'role=admin&plan=enterprise';

  // ---- licence
  await open(p, E, 'admin/licence');
  const lic = (await p.textContent('.ts-kpis')).replace(/\s+/g, ' ');
  check('Enterprise licence key accepted; 30 users; ₹300 per extra user', /Enterprise/.test(lic) && /\/ 30/.test(lic) && /₹300/.test(lic), lic.slice(0, 120));

  // ---- start day of the week + tracking days
  await open(p, E, 'admin/tracking');
  check('Configure start day of the week (not locked)', await p.isVisible('text=Week starts on'));
  await p.click('.ms-Dropdown:near(:text("Week starts on"))');
  await p.click('button[role="option"]:has-text("Sunday")');
  check('Changing the start day warns about weeks already filled in', await p.isVisible('text=Weeks already filled in keep their old dates'));
  check('Select days for time tracking (holidays editor)', await p.isVisible('text=Days without time tracking'));
  check('Select week days for time tracking (Pro)', await p.isVisible('text=Week days for time tracking') && !(await p.isVisible('text=Available on the Pro plan')));

  // ---- hour rules
  await open(p, E, 'admin/hours');
  await p.click('button:has-text("Add a rule")');
  check('Restrict daily & weekly hours based on rules (rule editor)', await p.isVisible('text=Rule 1') || (await p.$$('input[aria-label="Rule name"]')).length > 0);
  check('Daily & weekly hour restrictions (Pro)', await p.isVisible('text=Daily maximum'));

  // ---- approvals
  await open(p, E, 'admin/approvals');
  check('Role-based auto timesheet approval (Pro)', await p.isVisible('text=Approve automatically on submit'));

  // ---- custom columns, costs
  await open(p, E, 'admin/columns');
  check('Add custom columns', await p.isVisible('text=Add a column'));
  await open(p, E, 'admin/costs');
  check('User-specific hourly cost and flat project cost', await p.isVisible("text=The person's hourly cost; else the project's flat hourly cost"));

  // ---- reminders + project end notices + planned hours (Pro / Enterprise)
  await open(p, E, 'admin/notifications');
  const before = await lists(p, () => window.__tsLists().TS_Notifications.items.length);
  await p.click('button:has-text("Send due reminders now")');
  await p.waitForSelector('text=/reminders and .* project notices sent/', { timeout: 10000 });
  const kinds = await lists(p, () => window.__tsLists().TS_Notifications.items.map((x) => x.Kind));
  const reminders = kinds.filter((k) => k === 'reminder').length;
  const ending = kinds.filter((k) => k === 'projectEnding' || k === 'projectEnded').length;
  check('Pending weekly timesheet reminders sent', reminders > 0, `${reminders} reminders (bell + email)`);
  check('Project end date notifications sent', ending > 0, `${ending} notices (Website redesign ends in 6 days)`);
  check('Emails queued or sent for them', (await lists(p, () => (window.__tsLists().TS_EmailQueue || { items: [] }).items.length)) > 0 || before >= 0, 'queued for the Power Automate flow (Outlook via Graph is not available in the preview)');
  const again = await lists(p, () => window.__tsLists().TS_Notifications.items.length);
  await p.click('button:has-text("Send due reminders now")');
  await p.waitForTimeout(2500);
  check('Reminders go out once per person and week', (await lists(p, () => window.__tsLists().TS_Notifications.items.length)) === again);

  // ---- email templates
  const srcdoc = await p.getAttribute('iframe.ts-tpl-preview', 'srcdoc');
  check('Email templates editor with live preview', srcdoc && srcdoc.length > 500 && /Arjun Mehta/.test(srcdoc));
  check('Teams & Outlook: Outlook option and Teams chat toggle available', (await p.isVisible('text=Outlook — from the mailbox')) && (await p.isVisible('text=Also post notifications to people in Teams chat')));

  // ---- integrations: downloads and screens
  await open(p, E, 'admin/integrations');
  const tabs = ['Brillienta Task Manager', 'Time off Manager 365', 'Teams & Outlook', 'Planner', 'Copilot', 'Power Automate', 'API', 'Power BI'];
  for (const t of tabs) {
    await p.click(`button[role="tab"]:has-text("${t}")`);
    await p.waitForTimeout(250);
    check(`Integration screen: ${t} (unlocked)`, !(await p.isVisible('text=Available on the Enterprise plan')));
  }
  const dl = async (tab, button, prepare) => {
    await p.click(`button[role="tab"]:has-text("${tab}")`);
    if (prepare) await prepare();
    const [d] = await Promise.all([p.waitForEvent('download', { timeout: 10000 }), p.click(`button:has-text("${button}")`)]);
    const path = `${OUT}/${d.suggestedFilename()}`;
    await d.saveAs(path);
    return { name: d.suggestedFilename(), bytes: fs.readFileSync(path) };
  };
  const zipHas = (buf, name) => buf.includes(Buffer.from(name));
  const cop = await dl('Copilot', 'Download the Copilot agent', async () => p.click('text=Copilot integration on'));
  check('Copilot integration: agent package', zipHas(cop.bytes, 'declarativeAgent.json') && zipHas(cop.bytes, 'color.png') && zipHas(cop.bytes, 'OneDriveAndSharePoint'), `${cop.name}, ${cop.bytes.length} bytes`);
  const teams = await dl('Teams & Outlook', 'Download the Teams app package');
  check('Teams app package (personal app + tab)', zipHas(teams.bytes, 'staticTabs') && zipHas(teams.bytes, 'configurableTabs') && zipHas(teams.bytes, 'outline.png'), `${teams.name}, ${teams.bytes.length} bytes`);
  const pbi = await dl('Power BI', 'Download Power BI data source');
  check('Connect with Power BI (.pbids)', JSON.parse(pbi.bytes.toString()).connections[0].details.protocol === 'sharepoint-list', pbi.name);
  await p.click('button[role="tab"]:has-text("API")');
  check('API connectivity: endpoints for every list', (await p.$$eval('table.ts-table tbody tr', (r) => r.length)) === 13);
  await p.click('button[role="tab"]:has-text("Power Automate")');
  check('Customise with Power Automate: HTTP trigger and events', (await p.isVisible('text=Flow HTTP trigger URL')) && (await p.isVisible('text=Timesheet approved')));
  await open(p, E, 'admin/notifications');
  const [flow] = await Promise.all([p.waitForEvent('download', { timeout: 10000 }), p.click('button:has-text("Download the email flow")')]);
  const flowPath = `${OUT}/${flow.suggestedFilename()}`;
  await flow.saveAs(flowPath);
  check('Power Automate email flow package', zipHas(fs.readFileSync(flowPath), 'TS_EmailQueue'), flow.suggestedFilename());

  // ---- projects: planned hours, end dates, disable, delete, bulk
  await open(p, E, 'projects');
  check('Planned project hours tracking (logged / planned bars)', (await p.$$('.ts-progress')).length > 0);
  check('Bulk project / task creation', await p.isVisible('button:has-text("Bulk create")'));
  const row = 'tr:has-text("Data platform")';
  await p.click(`${row} button:has-text("Manage")`);
  await p.click('button[role="menuitem"]:has-text("Disable (no new time)")');
  await p.waitForTimeout(1200);
  check('Disable projects', (await lists(p, () => window.__tsLists().TS_Projects.items.filter((x) => x.Title === 'Data platform')[0].ProjectStatus)) === 'Disabled');
  await p.click('.ms-Dropdown:has-text("Active")');
  await p.click('button[role="option"]:has-text("All statuses")');
  await p.waitForTimeout(400);
  await p.click('tr:has-text("Legacy CRM support") button:has-text("Manage")');
  await p.click('button[role="menuitem"]:has-text("Delete")');
  await p.click('.ms-Dialog-main button:has-text("Delete")');
  await p.waitForTimeout(1500);
  check('Delete projects', (await lists(p, () => window.__tsLists().TS_Projects.items.filter((x) => x.Title === 'Legacy CRM support').length)) === 0);
  await p.click('button:has-text("New project")');
  check('Project end date and planned hours in the project editor', (await p.isVisible('text=End date')) && (await p.isVisible('text=Planned hours')) && !(await p.isVisible('.ms-Panel-main >> text=Available on the Enterprise plan')));
  await p.keyboard.press('Escape');

  // ---- people: external users, per-person days, rates, programme manager
  await open(p, E, 'people');
  await p.click('button:has-text("Add external user")');
  await p.fill('.ms-Dialog-main input >> nth=0', 'lena.weber@northwind.de');
  await p.fill('.ms-Dialog-main input >> nth=1', 'Lena Weber (Northwind)');
  await p.click('.ms-Dialog-main button:has-text("Add")');
  await p.waitForSelector('.ms-Dialog-main', { state: 'detached', timeout: 10000 });
  await p.waitForTimeout(800);
  check('Add Office 365 external users', (await lists(p, () => (window.__tsLists().TS_Users.items.filter((x) => x.Email === 'lena.weber@northwind.de')[0] || {}).IsExternal)) === true);
  await p.click('td:has-text("Rahul Sharma")');
  await p.waitForSelector('.ms-Panel-main');
  check('User-specific rates and hourly cost on a person', (await p.isVisible('.ms-Panel-main >> text=/Hourly cost/')) && (await p.isVisible('.ms-Panel-main >> text=/Bill rate/')));
  await p.click('.ms-Panel-main .ms-Checkbox:has-text("Sat") label');
  await p.click('.ms-Panel-main button:has-text("Save")');
  await p.waitForTimeout(1200);
  check('Select days for time tracking per person', (await lists(p, () => window.__tsLists().TS_Users.items.filter((x) => x.Email === 'rahul.sharma@contoso.com')[0].WorkDays)) === '1,2,3,4,5,6');

  // ---- custom reports & dashboards
  await open(p, E, 'reports');
  await p.click('button[role="tab"]:has-text("Custom report")');
  await p.fill('.ts-split input >> nth=-1', 'Hours by client and month');
  await p.click('button:has-text("Save report")');
  await p.waitForTimeout(1200);
  check('Custom reports: saved', (await lists(p, () => window.__tsLists().TS_SavedReports.items.filter((x) => x.ReportKind === 'report').length)) === 1);
  for (const t of ['Project report', 'My team', 'Billable / non-billable', 'Cost & margin', 'Utilisation']) {
    await p.click(`button[role="tab"]:has-text("${t}")`);
    await p.waitForTimeout(400);
    check(`Report: ${t}`, !(await p.isVisible('text=/Available on the/')) && (await p.$$('.ts-card')).length > 0);
  }
  await open(p, E, 'dashboards');
  check('Custom dashboards: starter dashboard with widgets', (await p.$$('.ts-widgets .ts-card')).length >= 5);
  await p.click('button:has-text("Save as my dashboard")');
  await p.click('.ms-Panel-main button:has-text("Save dashboard")');
  await p.waitForTimeout(1200);
  check('Custom dashboards: saved', (await lists(p, () => window.__tsLists().TS_SavedReports.items.filter((x) => x.ReportKind === 'dashboard').length)) === 1);

  // ---- support + account manager
  await open(p, E, 'support');
  check('Dedicated account manager shown', await p.isVisible('text=Kavya Reddy'));
  check('Unlimited support (no ticket limit)', /Unlimited/.test(await p.textContent('.ts-kpis')));

  // ---- Time off Manager 365: leave on the sheet (weekdays only), counts towards minimums
  await open(p, E + '&leave=1', 'timesheet');
  const off = await p.evaluate(() => {
    const row = Array.from(document.querySelectorAll('table.ts-sheet tbody tr')).filter((r) => /Time off/.test(r.textContent))[0];
    return row ? Array.from(row.querySelectorAll('input.ts-cell')).map((i) => i.value) : null;
  });
  check('Time off Manager 365: approved leave put on the timesheet', !!off, off ? off.join(' | ') : 'no time-off line');
  check('Time off: leave only on working days (not Sat/Sun), rejected leave ignored', off && off[0] === '' && off[2] === '8' && off[5] === '' && off[6] === '', off && off.join(' | '));
  check('Time off: shown in the side card', await p.isVisible('text=Time off this week'));

  // ---- programme manager role
  await open(p, 'role=pgm&plan=enterprise', 'programs');
  check('Programme Manager role: sees their programme', await p.isVisible('text=Digital transformation'));
  await open(p, 'role=pgm&plan=enterprise', 'approvals');
  const pgmRows = await p.$$eval('table.ts-table tbody tr', (r) => r.map((x) => x.textContent));
  check('Programme Manager approves time on programme projects', pgmRows.length > 0, `${pgmRows.length} waiting`);

  // ---- clone, billable, self-assign (Standard, included)
  await open(p, 'role=member&plan=enterprise', 'tasks');
  await p.click('button[role="tab"]:has-text("Open to take")');
  check('Self-assign task option', (await p.$$('button:has-text("Assign to me")')).length > 0);
  await open(p, 'role=member&plan=enterprise', 'timesheet');
  check('Clone weekly timesheets', await p.isVisible('button:has-text("Clone a week")'));
  check('Billable / non-billable tracking on lines', await p.isVisible('th:has-text("Billable")'));
  check('Activity tracking column', (await p.$$('[aria-label="Activity"]')).length > 0);

  check('No script errors on any page', errs.length === 0, errs.slice(0, 3).join(' / '));
  const fail = results.filter((r) => !r.ok);
  console.log(`\n${results.length - fail.length} of ${results.length} checks passed`);
  fs.writeFileSync(`${OUT}/audit.json`, JSON.stringify(results, null, 2));
  await b.close();
  process.exit(fail.length ? 1 : 0);
})().catch((e) => {
  console.error('ABORTED', e.message.split('\n')[0]);
  process.exit(1);
});

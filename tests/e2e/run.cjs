#!/usr/bin/env node
// End-to-end tests: complete journeys through the real app (compiled lib/) running in Chromium against the preview's
// in-memory SharePoint. People sign in and out inside one session (window.__tsSignIn), so a week submitted by a member
// is the same week their manager rejects and the member resubmits.
//
//   npm run e2e                 build, then run every journey
//   node tests/e2e/run.cjs J3   run journeys whose id starts with J3
//
// Failures save a screenshot to tests/e2e/output/.

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..', '..');
const BASE = 'file://' + path.join(ROOT, 'preview-dist', 'index.html');
const OUT = path.join(__dirname, 'output');
fs.mkdirSync(OUT, { recursive: true });

const only = process.argv[2] || '';
const results = [];
let current = '';

// ------------------------------------------------------------------ helpers

class Step extends Error {}

function expect(cond, message) {
  if (!cond) throw new Step(message);
}

async function newPage(browser, viewport) {
  const ctx = await browser.newContext({ viewport: viewport || { width: 1440, height: 1000 }, acceptDownloads: true });
  const page = await ctx.newPage();
  page.errors = [];
  page.on('pageerror', (e) => page.errors.push(e.message));
  return page;
}

const h = (page) => ({
  async open(query, route) {
    await page.goto(`${BASE}?${query || ''}#/${route || 'home'}`);
    await page.waitForSelector('.ts-page, .ts-setup, .ts-center', { timeout: 20000 });
    // Past the loading spinner (the setup wizard and the lock screen are pages too).
    await page.waitForFunction(() => !document.querySelector('.ms-Spinner'), undefined, { timeout: 20000 });
    await page.waitForTimeout(1000);
  },
  async signIn(role, route) {
    await page.evaluate(([r, hash]) => window.__tsSignIn(r, hash), [role, '#/' + (route || 'home')]);
    await page.waitForSelector('.ts-page', { timeout: 20000 });
    await page.waitForTimeout(1300);
  },
  async go(route) {
    await page.evaluate((r) => (window.location.hash = '#/' + r), route);
    await page.waitForTimeout(1300);
  },
  lists(fn, arg) {
    return page.evaluate(fn, arg);
  },
  async toast(re, timeout) {
    const loc = page.locator('.ts-toasts').getByText(re).first();
    await loc.waitFor({ timeout: timeout || 10000 });
    return loc.textContent();
  },
  async pick(label, option, nth) {
    const boxes = page.locator(`[aria-label="${label}"]`);
    await boxes.nth(nth === undefined ? -1 : nth).click();
    await page.locator('button[role="option"]').filter({ hasText: option }).first().click();
  },
  async dropdownByLabel(label, option) {
    await page.locator('.ms-Dropdown-container').filter({ has: page.locator('label', { hasText: label }) }).locator('.ms-Dropdown').first().click();
    await page.locator('button[role="option"]').filter({ hasText: option }).first().click();
  },
  async fillHours(rowIndex, values) {
    const row = page.locator('table.ts-sheet tbody tr:not(.ts-sheet-total)').nth(rowIndex);
    const cells = row.locator('input.ts-cell');
    for (let i = 0; i < values.length; i++) {
      if (values[i] === null) continue;
      await cells.nth(i).fill(String(values[i]));
      await cells.nth(i).press('Tab');
    }
  },
  async click(text, scope) {
    await page.locator(`${scope || ''} button:has-text("${text}")`.trim()).first().click();
  },
  async dialogButton(text) {
    await page.locator('.ms-Dialog-main button').filter({ hasText: text }).first().click();
  },
  async total() {
    return (await page.locator('tr.ts-sheet-total td.ts-col-total').textContent()).trim();
  },
  status() {
    return page.locator('.ts-weekbar .ts-status').first().textContent();
  },
  sheetOf(email) {
    return page.evaluate((e) => window.__tsLists().BrillTS_Timesheets.items.filter((x) => x.UserEmail === e).sort((a, b) => b.WeekStart.localeCompare(a.WeekStart))[0], email);
  }
});

async function journey(id, title, browser, fn) {
  if (only && id.indexOf(only) !== 0) return;
  current = id;
  const page = await newPage(browser);
  const steps = [];
  const log = (s) => {
    steps.push(s);
    console.log(`    ✓ ${s}`);
  };
  console.log(`\n${id}  ${title}`);
  const started = Date.now();
  try {
    await fn(page, h(page), log);
    // List names: the app only ever touches BrillTS_ lists (plus Task Manager's TM_ and the leave list it is pointed at).
    const names = await page.evaluate(() => (window.__tsLists ? Object.keys(window.__tsLists()) : []));
    const known = ['Settings', 'Users', 'Programs', 'Projects', 'Tasks', 'Activities', 'Timesheets', 'TimeEntries', 'SupportTickets', 'Notifications', 'EmailQueue', 'AuditLog', 'SavedReports'].map((x) => 'BrillTS_' + x);
    const foreign = names.filter((n) => !/^(TM_|TOM_)/.test(n) && known.indexOf(n) < 0);
    expect(!foreign.length, 'lists outside the BrillTS_ prefix were used: ' + foreign.join(', '));
    expect(!page.errors.length, 'script errors: ' + page.errors.slice(0, 2).join(' / '));
    results.push({ id, title, ok: true, steps, ms: Date.now() - started });
    console.log(`  PASS (${((Date.now() - started) / 1000).toFixed(1)} s)`);
  } catch (e) {
    const shot = path.join(OUT, `${id}-failure.png`);
    await page.screenshot({ path: shot }).catch(() => undefined);
    results.push({ id, title, ok: false, steps, error: e.message.split('\n')[0], shot });
    console.log(`  FAIL: ${e.message.split('\n')[0]}\n  screenshot: ${shot}`);
  } finally {
    await page.context().close();
  }
}

const MEMBER = 'rahul.sharma@contoso.com';

// ------------------------------------------------------------------ journeys

async function main() {
  const browser = await chromium.launch();

  await journey('J01', 'First run: a site owner sets up the app with the setup wizard', browser, async (page, t, log) => {
    await t.open('fresh=1&role=admin');
    expect(await page.isVisible('text=Set up Brillienta Timesheet'), 'setup wizard not shown');
    log('setup wizard shown on a site without the lists');
    for (let i = 0; i < 3; i++) await t.click('Next');
    await t.click('Create the timesheet');
    await page.waitForSelector('.ts-nav', { timeout: 30000 });
    await page.waitForTimeout(1500);
    const lists = await t.lists(() => Object.keys(window.__tsLists()));
    const expected = ['Settings', 'Users', 'Programs', 'Projects', 'Tasks', 'Activities', 'Timesheets', 'TimeEntries', 'SupportTickets', 'Notifications', 'EmailQueue', 'AuditLog', 'SavedReports'].map((x) => 'BrillTS_' + x);
    const missing = expected.filter((x) => lists.indexOf(x) < 0);
    expect(!missing.length, 'lists missing: ' + missing.join(', '));
    log(`all 13 lists created with the BrillTS_ prefix (${expected.slice(0, 3).join(', ')}, …)`);
    const counts = await t.lists(() => ({ a: window.__tsLists().BrillTS_Activities.items.length, u: window.__tsLists().BrillTS_Users.items.map((x) => x.UserRole + ':' + x.Email), p: window.__tsLists().BrillTS_Projects.items.length, t: window.__tsLists().BrillTS_Tasks.items.length }));
    expect(counts.a === 9 && counts.p === 2 && counts.t === 4, JSON.stringify(counts));
    expect(counts.u.length === 1 && /^admin:/.test(counts.u[0]), 'installer not added as administrator');
    log('9 activities, 2 sample projects, 4 tasks; the installer is an administrator');
    expect(await page.isVisible('text=Trial (every Enterprise feature): 14 days left'), 'trial banner');
    log('14-day Enterprise trial started');
  });

  await journey('J02', 'Administrator builds a project: form, task, bulk tasks, bulk activities', browser, async (page, t, log) => {
    await t.open('role=admin', 'projects');
    await t.click('New project');
    await page.waitForSelector('.ms-Panel-main');
    await page.getByLabel('Project name').fill('ERP rollout');
    await page.getByLabel('Code').fill('ERP-07');
    await page.getByLabel('Client').fill('Woodgrove Bank');
    await page.getByLabel('Planned hours').fill('600');
    await page.locator('.ms-Panel-main').getByLabel(/Flat hourly cost/).fill('1500');
    await page.locator('.ms-Panel-main').getByLabel(/Hourly bill rate/).fill('3200');
    const members = page.locator('.ms-Panel-main .ts-picker').nth(1).locator('input');
    for (const who of ['Rahul', 'Sara']) {
      await members.fill(who);
      await page.waitForTimeout(300);
      await members.press('Enter');
    }
    await page.locator('.ms-Panel-main').getByLabel('PO number').fill('PO-9001');
    await t.click('Save project', '.ms-Panel-main');
    await page.waitForSelector('text=Linked to Task Manager >> nth=0', { timeout: 1500 }).catch(() => undefined);
    await page.waitForSelector('h1:has-text("ERP rollout")', { timeout: 10000 });
    const p = await t.lists(() => window.__tsLists().BrillTS_Projects.items.filter((x) => x.Title === 'ERP rollout')[0]);
    expect(p && p.ProjectCode === 'ERP-07' && p.PlannedHours === 600 && p.HourlyCost === 1500 && /rahul/.test(p.MemberEmails) && /sara/.test(p.MemberEmails) && /PO-9001/.test(p.CustomValues), JSON.stringify(p));
    log('project created with code, client, members, planned hours, cost, rate and a custom column');
    await page.locator('button[role="tab"]:has-text("Tasks")').click();
    await t.click('New task');
    await page.locator('.ms-Panel-main').getByLabel('Task', { exact: true }).fill('Data migration');
    const assignee = page.locator('.ms-Panel-main .ts-picker input');
    await assignee.fill('Rahul');
    await page.waitForTimeout(300);
    await assignee.press('Enter');
    await page.locator('.ms-Panel-main').getByLabel('Planned hours').fill('120');
    await t.click('Save', '.ms-Panel-main');
    await page.waitForTimeout(1500);
    const task = await t.lists(() => window.__tsLists().BrillTS_Tasks.items.filter((x) => x.Title === 'Data migration')[0]);
    expect(task && /rahul/.test(task.AssigneeEmails), 'task not created / not assigned');
    const note = await t.lists(() => window.__tsLists().BrillTS_Notifications.items.filter((x) => x.Kind === 'taskAssigned' && x.Recipient === 'rahul.sharma@contoso.com').length);
    expect(note === 1, 'assignee not notified');
    log('task created, assigned to Rahul, Rahul notified');
    await t.click('Bulk create tasks');
    await page.locator('.ms-Dialog-main textarea').fill('Project\tTask\tAssignees\tBillable\tPlanned hours\nERP-07\tUser training\t\tyes\t40\nERP-07\tCut-over\tsara.khan@contoso.com\tyes\t24\nNOPE\tBad row\t\t\t');
    await page.waitForTimeout(400);
    expect(await page.isVisible('text=2 of 3 rows ready'), 'bulk validation');
    log('bulk paste validated: 2 of 3 rows ready (unknown project rejected)');
    await t.dialogButton('Create 2 tasks');
    await page.waitForTimeout(1500);
    const tasks = await t.lists(() => window.__tsLists().BrillTS_Tasks.items.filter((x) => ['User training', 'Cut-over'].indexOf(x.Title) >= 0).length);
    expect(tasks === 2, `bulk tasks: ${tasks}`);
    log('2 tasks created in bulk');
    await t.go('activities');
    await t.click('Bulk create');
    await page.locator('.ms-Dialog-main textarea').fill('Activity\tBillable\nTravel\tno\nCode review\tyes');
    await page.waitForTimeout(300);
    await t.dialogButton('Create 2 activities');
    await page.waitForTimeout(1500);
    expect((await t.lists(() => window.__tsLists().BrillTS_Activities.items.filter((x) => x.Title === 'Travel' && x.IsBillable === false).length)) === 1, 'bulk activities');
    log('activities created in bulk (Travel non-billable)');
  });

  await journey('J03', 'Weekly cycle: member submits, manager sends back, member corrects, manager approves', browser, async (page, t, log) => {
    await t.open('role=member', 'timesheet');
    await t.pick('Project', 'Website redesign');
    await t.pick('Task', 'Front-end build');
    await t.pick('Activity', 'Development');
    await t.fillHours(0, [9, 8, 8, 8, 8]);
    expect((await t.total()) === '41', 'total ' + (await t.total()));
    await t.click('Add line');
    await t.fillHours(1, [1]);
    await t.click('Submit');
    await page.waitForTimeout(400);
    expect(await page.isVisible('text=Every line with hours needs a project.'), 'missing-project error not shown');
    expect(!(await page.isVisible('.ms-Dialog-main')), 'submit dialog opened despite errors');
    log('submit blocked: a line with hours has no project');
    await page.locator('table.ts-sheet tbody tr:not(.ts-sheet-total)').nth(1).locator('button[aria-label="More actions"]').click();
    await page.locator('button[role="menuitem"]:has-text("Remove line")').click();
    await page.locator('table.ts-sheet tbody tr').first().locator('button[aria-label="Notes"]').click();
    await page.locator('input[aria-label^="Note for"]').first().fill('Header and navigation');
    await t.click('Submit');
    await page.locator('.ms-Dialog-main textarea').fill('Busy week on the redesign');
    await t.dialogButton('Submit');
    await t.toast(/Submitted for approval/);
    let s = await t.sheetOf(MEMBER);
    expect(s.SheetStatus === 'Submitted' && s.TotalHours === 41, JSON.stringify(s));
    const approvers = await t.lists((id) => window.__tsLists().BrillTS_Notifications.items.filter((x) => x.Kind === 'sheetSubmitted' && x.Created > id).map((x) => x.Recipient), '');
    expect(approvers.indexOf('meera.iyer@contoso.com') >= 0 && approvers.indexOf('arjun.mehta@contoso.com') >= 0, 'approvers: ' + approvers);
    log('41 h submitted with a note and a comment; manager and project manager notified');
    expect(await page.isVisible('button:has-text("Recall")'), 'recall offered');
    await t.signIn('manager', 'approvals');
    const badge = (await page.locator('.ts-nav-item:has-text("Approvals") .ts-nav-count').textContent().catch(() => '0')).trim();
    expect(Number(badge) >= 1, 'approvals badge ' + badge);
    await page.locator('table.ts-table tbody tr').filter({ hasText: 'Rahul Sharma' }).filter({ hasText: '41' }).first().click();
    await page.waitForSelector('.ms-Panel-main');
    expect(await page.isVisible('.ms-Panel-main >> text=Header and navigation'), 'note not visible to approver');
    log(`manager sees ${badge} waiting, opens Rahul's week with the note`);
    await page.locator('.ms-Panel-main button:has-text("Send back")').click();
    await page.locator('.ms-Dialog-main textarea').fill('Monday should be 8 hours');
    await t.dialogButton('Send back');
    await t.toast(/sent back/);
    s = await t.sheetOf(MEMBER);
    expect(s.SheetStatus === 'Rejected' && s.DecidedBy === 'Meera Iyer', JSON.stringify(s));
    log('manager sent it back with a reason');
    await t.signIn('member', 'timesheet');
    const unread = (await page.locator('.ts-bell-count').textContent()).trim();
    expect(Number(unread) >= 1, 'no unread notification');
    expect(await page.isVisible('text=Monday should be 8 hours'), 'rejection reason not shown');
    expect((await t.status()) === 'Rejected', 'status ' + (await t.status()));
    log(`member sees the reason and ${unread} unread notification(s)`);
    await t.fillHours(0, [8]);
    await t.click('Resubmit');
    await t.dialogButton('Submit');
    await t.toast(/Submitted for approval/);
    expect((await t.sheetOf(MEMBER)).TotalHours === 40, 'resubmitted total');
    log('member corrected Monday and resubmitted 40 h');
    await t.signIn('manager', 'approvals');
    await page.locator('table.ts-table tbody tr').filter({ hasText: 'Rahul Sharma' }).filter({ hasText: 'just now' }).first().click();
    await page.locator('.ms-Panel-main button:has-text("Approve")').click();
    await t.toast(/approved/);
    s = await t.sheetOf(MEMBER);
    expect(s.SheetStatus === 'Approved', s.SheetStatus);
    const rates = await t.lists((sid) => window.__tsLists().BrillTS_TimeEntries.items.filter((x) => x.SheetId === sid).map((x) => x.CostRate + '/' + x.BillRate), s.Id);
    expect(rates.length && rates.every((r) => r === '1400/2800'), 'rates captured: ' + rates);
    log('manager approved; Rahul’s cost 1400 and bill rate 2800 captured on the lines');
    await t.signIn('member', 'timesheet');
    expect((await t.status()) === 'Approved', 'member status');
    const history = await page.locator('.ts-timeline-item').count();
    expect(history === 4, 'history items ' + history);
    expect(!(await page.isVisible('input.ts-cell:not([disabled])')), 'approved week still editable');
    log('member sees Approved, read-only, with the 4-step history');
  });

  await journey('J04', 'Role-based auto approval: a programme manager’s week is approved on submit', browser, async (page, t, log) => {
    await t.open('role=pgm', 'timesheet');
    await t.pick('Project', 'Mobile app');
    await t.fillHours(0, [8, 8, 8, 8, 8]);
    await t.click('Submit');
    await t.dialogButton('Submit');
    await t.toast(/approved automatically/);
    const s = await t.sheetOf('vikram.rao@contoso.com');
    expect(s.SheetStatus === 'Approved' && s.AutoApproved === true && s.DecidedBy === 'Auto-approval', JSON.stringify(s));
    expect(await page.isVisible('text=Auto-approved'), 'auto-approved chip');
    log('submitted and approved automatically (programme managers are auto-approved)');
  });

  await journey('J05', 'Hour restrictions: an administrator blocks weeks over 45 h; a member cannot submit 50 h', browser, async (page, t, log) => {
    await t.open('role=admin', 'admin/hours');
    await page.getByLabel('Weekly maximum').first().fill('45');
    await page.locator('label:has-text("Block submitting until it is fixed")').click();
    await t.click('Save');
    await t.toast(/Settings saved/);
    log('weekly maximum 45 h, enforced');
    await t.signIn('member', 'timesheet');
    // Rahul's draft for this week already has two lines: 50 h on the first, none on the second.
    await t.fillHours(0, [10, 10, 10, 10, 10]);
    await t.fillHours(1, [0, 0, 0, 0, 0]);
    await t.click('Submit');
    await page.waitForTimeout(400);
    expect(await page.isVisible('.ts-issues .is-error:has-text("more than the weekly maximum of 45 h")'), 'weekly max error not shown');
    expect(!(await page.isVisible('.ms-Dialog-main')), 'submit allowed');
    log('50 h blocked with the reason shown');
    await t.fillHours(0, [9, 9, 9, 9, 9]);
    await t.click('Submit');
    await page.waitForSelector('.ms-Dialog-main');
    log('45 h can be submitted');
  });

  await journey('J06', 'Reports and export reflect approved time, cost and margin', browser, async (page, t, log) => {
    await t.open('role=admin', 'reports');
    await t.dropdownByLabel('Period', 'Last 90 days');
    await page.waitForTimeout(1500);
    const hoursText = await page.locator('.ts-kpi').filter({ has: page.locator('.ts-kpi-label', { hasText: /^Hours$/ }) }).locator('.ts-kpi-value').first().textContent();
    const hours = Number(hoursText.replace(/[^\d.]/g, ''));
    const facts = await t.lists(() => window.__tsLists().BrillTS_TimeEntries.items.reduce((s, e) => s + Object.values(JSON.parse(e.DayHours || '{}')).reduce((a, b) => a + b, 0), 0));
    expect(hours > 0 && Math.abs(hours - facts) < 0.01, `report ${hours} h vs data ${facts} h`);
    log(`project report total ${hours} h matches the timesheet data`);
    await page.locator('button[role="tab"]:has-text("Cost & margin")').click();
    await page.waitForTimeout(500);
    expect(await page.isVisible('text=Margin'), 'margin');
    log('cost & margin report shown');
    await page.locator('button[role="tab"]:has-text("Project report")').click();
    await page.waitForTimeout(500);
    const [d] = await Promise.all([page.waitForEvent('download'), page.locator('.ts-card button:has-text("Export")').first().click()]);
    const file = path.join(OUT, d.suggestedFilename());
    await d.saveAs(file);
    const bytes = fs.readFileSync(file);
    expect(bytes[0] === 0x50 && bytes[1] === 0x4b && bytes.includes(Buffer.from('xl/worksheets/sheet2.xml')), 'not an xlsx with 2 sheets');
    log(`Excel export: ${d.suggestedFilename()} (${bytes.length} bytes, summary + details sheets)`);
  });

  await journey('J07', 'Self-assign: a member takes an open task and logs time on it', browser, async (page, t, log) => {
    await t.open('role=member', 'tasks');
    await page.locator('button[role="tab"]:has-text("Open to take")').click();
    await page.locator('tr:has-text("Accessibility review") button:has-text("Assign to me")').click();
    await t.toast(/is now yours/);
    expect(/rahul/.test(await t.lists(() => window.__tsLists().BrillTS_Tasks.items.filter((x) => x.Title === 'Accessibility review')[0].AssigneeEmails)), 'not assigned');
    log('task assigned to the member');
    await t.go('timesheet');
    await t.pick('Project', 'Website redesign');
    await t.pick('Task', 'Accessibility review');
    log('task offered on the timesheet');
  });

  await journey('J08', 'Clone a week, save, submit, recall', browser, async (page, t, log) => {
    // The member's current week is empty: clone last week (approved or submitted) into it.
    await t.open('role=member', 'timesheet');
    await t.click('Clone a week');
    await t.dialogButton('Clone');
    await page.waitForTimeout(500);
    const total = await t.total();
    expect(Number(total) > 0, 'nothing cloned');
    log(`last week's lines and ${total} h cloned`);
    await t.click('Save draft');
    await t.toast(/Saved/);
    await t.click('Submit');
    await t.dialogButton('Submit');
    await t.toast(/Submitted/);
    await t.click('Recall');
    await t.dialogButton('Recall');
    await t.toast(/Back to draft/);
    expect((await t.status()) === 'Draft', 'not draft');
    log('saved, submitted and recalled back to draft');
  });

  await journey('J09', 'An administrator reopens an approved week', browser, async (page, t, log) => {
    await t.open('role=admin', 'approvals');
    await page.locator('button[role="tab"]:has-text("Decided")').click();
    await page.waitForTimeout(1500);
    await page.locator('table.ts-table tbody tr').filter({ hasText: 'Sara Khan' }).first().click();
    const head = await page.locator('.ms-Panel-headerText').textContent();
    await page.locator('.ms-Panel-main button:has-text("Reopen")').click();
    await page.locator('.ms-Dialog-main textarea').fill('Client asked for task detail');
    await t.dialogButton('Reopen');
    await t.toast(/Reopened/);
    const week = /· (.*)$/.exec(head)[1];
    log(`Sara's approved week (${week}) reopened`);
    await t.signIn('sara', 'history');
    expect(await page.isVisible('text=Draft'), 'not draft for Sara');
    log('Sara sees it as a draft again');
  });

  await journey('J10', 'Licence: the trial has ended; an administrator activates a Pro key', browser, async (page, t, log) => {
    await t.open('role=admin&installed=30');
    expect(await page.isVisible('text=The 14-day trial has ended'), 'not locked');
    expect(!(await page.isVisible('.ts-nav')), 'app usable after the trial');
    log('app locked after the 14-day trial; data kept');
    const key = await page.evaluate(() => window.__tsPreviewKey('pro'));
    await page.locator('textarea').first().fill('ED1.not-a-timesheet-key.x');
    await t.click('Activate');
    await page.waitForSelector('text=This is an Employee Directory licence key');
    log('an Employee Directory key is refused');
    await page.locator('textarea').first().fill(key);
    await t.click('Activate');
    await page.waitForSelector('.ts-nav', { timeout: 15000 });
    await t.go('admin/licence');
    expect(/Pro/.test(await page.locator('.ts-kpis').textContent()), 'plan not Pro');
    log('Pro key activated; app open on the Pro plan');
  });

  await journey('J11', 'Standard plan: 4 support tickets a year', browser, async (page, t, log) => {
    await t.open('role=admin&plan=standard', 'support');
    for (let i = 1; i <= 4; i++) {
      await page.getByLabel('Subject').fill(`Question ${i}`);
      await page.getByLabel('Details').fill('Details of the question');
      await t.click('Send ticket');
      await t.toast(/Ticket sent/);
      await page.waitForTimeout(600);
    }
    expect(await page.locator('button:has-text("Send ticket")').isDisabled(), 'fifth ticket allowed');
    expect(await page.isVisible('text=(0 left in'), 'quota text');
    const queued = await t.lists(() => window.__tsLists().BrillTS_EmailQueue.items.filter((x) => /support@brillienta.com/.test(x.MailTo)).length);
    expect(queued === 4, 'tickets emailed: ' + queued);
    log('4 tickets sent to Brillienta support; the 5th is blocked');
  });

  await journey('J12', 'Notifications: unread count, open, mark all read', browser, async (page, t, log) => {
    await t.open('role=admin');
    const n = Number((await page.locator('.ts-bell-count').textContent()).trim());
    expect(n >= 1, 'no unread');
    await page.locator('.ts-bell button').click();
    await page.locator('.ts-bell-menu button:has-text("Mark all read")').click();
    await page.waitForTimeout(800);
    expect(!(await page.isVisible('.ts-bell-count')), 'still unread');
    log(`${n} unread → marked read`);
  });

  await journey('J13', 'Access: members cannot reach management pages; Standard locks Pro features', browser, async (page, t, log) => {
    await t.open('role=member', 'admin/general');
    expect(!(await page.isVisible('text=Settings for everyone')), 'member reached admin');
    for (const r of ['approvals', 'people', 'reports', 'projects']) {
      await t.go(r);
      expect(!(await page.isVisible(`.ts-crumb strong:text-is("${r === 'people' ? 'People & rates' : r[0].toUpperCase() + r.slice(1)}")`)), `member reached ${r}`);
    }
    expect((await page.locator('.ts-nav-item').count()) === 4, 'member nav items');
    log('member sees only Overview, My timesheet, My timesheets, My tasks');
    await t.open('role=admin&plan=standard', 'admin/approvals');
    expect(await page.isVisible('text=Available on the Pro plan'), 'auto approval not locked on Standard');
    log('Standard plan locks auto approval (Pro)');
  });

  await journey('J14', 'Phone width: the app fits and the week grid scrolls inside its frame', browser, async (page, t, log) => {
    await page.setViewportSize({ width: 390, height: 844 });
    await t.open('role=member', 'timesheet');
    const fit = await page.evaluate(() => ({ doc: document.documentElement.scrollWidth, win: window.innerWidth, grid: (document.querySelector('.ts-sheet-wrap') || {}).scrollWidth }));
    expect(fit.doc <= fit.win + 1, `page overflows: ${fit.doc} > ${fit.win}`);
    expect(fit.grid > 390, 'grid not scrollable');
    await page.locator('.ts-menu-btn').click();
    await page.waitForTimeout(300);
    expect(await page.locator('.ts-nav.is-open').isVisible(), 'menu does not open');
    log('no page overflow at 390 px; grid scrolls; menu opens');
  });

  await journey('J15', 'Integrations: Time off Manager 365 and Brillienta Task Manager', browser, async (page, t, log) => {
    await t.open('role=member&leave=1', 'timesheet');
    const off = await page.evaluate(() => {
      const row = Array.from(document.querySelectorAll('table.ts-sheet tbody tr')).filter((r) => /Time off/.test(r.textContent))[0];
      return row ? Array.from(row.querySelectorAll('input.ts-cell')).map((i) => i.value).join(',') : '';
    });
    expect(off === ',,8,8,8,,', 'leave line: ' + off);
    log('approved leave filled Wed–Fri (8 h), not the weekend; the rejected request ignored');
    await t.open('role=admin&tm=1', 'projects');
    await t.toast(/Task Manager:/, 15000);
    expect((await t.lists(() => window.__tsLists().BrillTS_Projects.items.filter((x) => x.TmProjectId > 0).length)) === 2, 'projects not synced');
    log('Task Manager projects and tasks synced on open');
    await t.signIn('member', 'timesheet');
    await t.click('Import from Task Manager');
    await t.toast(/imported from Task Manager/);
    await t.click('Add line');
    await t.pick('Project', 'Customer portal');
    await t.pick('Task', 'Sign-in and accounts');
    const last = (await page.locator('table.ts-sheet tbody tr:not(.ts-sheet-total)').count()) - 1;
    await t.fillHours(last, [null, 4, 4]);
    await t.click('Submit');
    await t.dialogButton('Submit');
    await t.toast(/Submitted/);
    await t.signIn('manager', 'approvals');
    await page.locator('table.ts-table tbody tr').filter({ hasText: 'Rahul Sharma' }).filter({ hasText: 'just now' }).first().click();
    const spentBefore = await t.lists(() => window.__tsLists().TM_Tasks.items.filter((x) => x.TaskNumber === 'TM-101')[0].SpentMinutes);
    await page.locator('.ms-Panel-main button:has-text("Approve")').click();
    await t.toast(/approved/);
    await page.waitForTimeout(800);
    const spentAfter = await t.lists(() => window.__tsLists().TM_Tasks.items.filter((x) => x.TaskNumber === 'TM-101')[0].SpentMinutes);
    // 8 h on the linked task (4 + 4); Rahul's other lines are not on Task Manager tasks; the imported 2 h is not sent back.
    expect(spentAfter - spentBefore === 480, `spent ${spentBefore} → ${spentAfter}`);
    log('2 h imported from Task Manager (not sent back); 8 h approved and written to Task Manager (+480 min)');
  });

  await journey('J16', 'Security: BrillTS groups created, kept in step with roles, lists granted to them', browser, async (page, t, log) => {
    await t.open('role=admin', 'admin/security');
    expect(await page.isVisible('text=BrillTS Administrators') && (await page.isVisible('text=BrillTS Approvers')) && (await page.isVisible('text=BrillTS External Users')), 'group names not shown');
    await t.click('Apply security');
    await t.toast(/Security applied/, 20000);
    const groups = await t.lists(() => Object.values(window.__tsGroups || {}).map((g) => ({ id: g.Id, title: g.Title, members: g.users.map((u) => u.Email).sort(), open: g.OnlyAllowMembersViewMembership === false })));
    const by = (title) => groups.filter((g) => g.title === title)[0];
    expect(groups.length === 3 && groups.every((g) => /^BrillTS /.test(g.title)), 'groups: ' + groups.map((g) => g.title).join(', '));
    log('created: ' + groups.map((g) => g.title).join(', '));
    expect(JSON.stringify(by('BrillTS Administrators').members) === JSON.stringify(['priya.nair@contoso.com']), 'admins: ' + by('BrillTS Administrators').members);
    expect(JSON.stringify(by('BrillTS Approvers').members) === JSON.stringify(['arjun.mehta@contoso.com', 'meera.iyer@contoso.com', 'vikram.rao@contoso.com']), 'approvers: ' + by('BrillTS Approvers').members);
    expect(JSON.stringify(by('BrillTS External Users').members) === JSON.stringify(['ana.lopez@fabrikam.com']), 'externals: ' + by('BrillTS External Users').members);
    expect(groups.every((g) => g.open), 'membership not visible to the site');
    log('members from the roles: 1 administrator, 3 approvers (programme, project and team managers), 1 external user');
    const EDIT = 1073741830;
    const CONTRIBUTE = 1073741827;
    const READ = 1073741826;
    const perms = await t.lists(() => {
      const out = {};
      Object.keys(window.__tsLists()).filter((k) => k.indexOf('BrillTS_') === 0).forEach((k) => {
        const l = window.__tsLists()[k];
        out[k] = { unique: !!l.__unique, roles: l.__roles || {}, read: (l.__props || {}).ReadSecurity };
      });
      return out;
    });
    const role = (list, title) => perms[list].roles[by(title).id];
    expect(Object.keys(perms).length === 13 && Object.values(perms).every((p) => p.unique), 'not every list has its own permissions');
    expect(role('BrillTS_TimeEntries', 'BrillTS Approvers') === EDIT && role('BrillTS_TimeEntries', 'BrillTS Administrators') === EDIT && role('BrillTS_TimeEntries', 'BrillTS External Users') === CONTRIBUTE, 'time entries grants ' + JSON.stringify(perms.BrillTS_TimeEntries.roles));
    expect(perms.BrillTS_TimeEntries.read === 2 && perms.BrillTS_Timesheets.read === 2, 'timesheets not limited to own rows');
    expect(role('BrillTS_Settings', 'BrillTS Approvers') === READ && role('BrillTS_Settings', 'BrillTS Administrators') === EDIT && role('BrillTS_Projects', 'BrillTS Approvers') === EDIT && role('BrillTS_Projects', 'BrillTS External Users') === READ, 'settings / projects grants');
    log('13 lists with their own permissions: timesheets own-rows-only; approvers and administrators edit all; guests contribute; settings read-only except administrators');
    await t.go('people');
    await page.locator('td:has-text("Tom Becker")').click();
    await t.dropdownByLabel('Role', 'Project Manager');
    await t.click('Save', '.ms-Panel-main');
    await t.toast(/Saved/);
    await page.waitForTimeout(1500);
    const approvers = await t.lists(() => Object.values(window.__tsGroups).filter((g) => g.Title === 'BrillTS Approvers')[0].users.map((u) => u.Email));
    expect(approvers.indexOf('tom.becker@contoso.com') >= 0, 'role change not reflected: ' + approvers);
    log('Tom made a Project Manager → added to BrillTS Approvers automatically');
    await page.locator('td:has-text("Ana Lopez")').click();
    await page.locator('.ms-Panel-main .ms-Toggle:has-text("Active") button').click();
    await t.click('Save', '.ms-Panel-main');
    await t.toast(/Saved/);
    await page.waitForTimeout(1500);
    const ext = await t.lists(() => Object.values(window.__tsGroups).filter((g) => g.Title === 'BrillTS External Users')[0].users.length);
    expect(ext === 0, 'deactivated guest still in BrillTS External Users');
    log('Ana deactivated → removed from BrillTS External Users automatically');
  });

  // ---------------------------------------------------------------- settings that change behaviour
  const L = 'window.__tsLists()';
  const saveAdmin = async (page, t) => {
    await page.locator('.ts-page button:has-text("Save")').last().click();
    await t.toast(/Settings saved/);
  };

  await journey('J17', 'General settings: organisation name and currency take effect', browser, async (page, t, log) => {
    await t.open('role=admin', 'admin/general');
    await page.getByLabel('Organisation name').fill('Contoso India');
    await t.dropdownByLabel('Currency', 'USD');
    await saveAdmin(page, t);
    expect((await page.locator('.ts-nav-workspace strong').textContent()) === 'Contoso India', 'organisation name not shown in the navigation');
    log('organisation name shown in the navigation');
    await t.go('reports');
    await page.locator('button[role="tab"]:has-text("Cost & margin")').click();
    await page.waitForTimeout(600);
    const cost = await page.locator('.ts-kpi').filter({ hasText: 'Cost' }).first().locator('.ts-kpi-value').textContent();
    expect(/^\$/.test(cost.trim()), 'cost not in USD: ' + cost);
    log(`reports use the new currency (${cost.trim()})`);
  });

  await journey('J18', 'Week start: weeks begin on Sunday after the change', browser, async (page, t, log) => {
    await t.open('role=admin', 'admin/tracking');
    await t.dropdownByLabel('Week starts on', 'Sunday');
    await saveAdmin(page, t);
    await t.signIn('member', 'timesheet');
    const first = (await page.locator('table.ts-sheet th.ts-day').first().textContent()).trim();
    expect(/^Sun/.test(first), 'first day column: ' + first);
    log('the timesheet grid now starts on Sunday');
  });

  await journey('J19', 'Required notes and activities block submitting until filled in', browser, async (page, t, log) => {
    await t.open('role=admin', 'admin/tracking');
    await page.locator('.ms-Toggle:has-text("A note is required on every day with hours") button').click();
    await page.locator('.ms-Toggle:has-text("An activity is required on every line") button').click();
    await saveAdmin(page, t);
    log('notes and activities made required');
    await t.signIn('member', 'timesheet');
    await t.click('Add line');
    await t.pick('Project', 'Website redesign');
    const last = (await page.locator('table.ts-sheet tbody tr:not(.ts-sheet-total)').count()) - 1;
    await t.fillHours(last, [1]);
    await t.click('Submit');
    await page.waitForTimeout(500);
    expect(await page.isVisible('.ts-issues .is-error:has-text("add a note")'), 'missing-note error');
    expect(await page.isVisible('.ts-issues .is-error:has-text("needs an activity")'), 'missing-activity error');
    expect(!(await page.isVisible('.ms-Dialog-main')), 'submit allowed');
    log('submit blocked: notes missing and a line without an activity');
  });

  await journey('J20', 'Required custom column on timesheet lines', browser, async (page, t, log) => {
    await t.open('role=admin', 'admin/columns');
    await page.locator('.ms-Checkbox:has-text("Required") label').first().click();
    await saveAdmin(page, t);
    await t.signIn('member', 'timesheet');
    await t.click('Add line');
    await t.pick('Project', 'Mobile app');
    const last = (await page.locator('table.ts-sheet tbody tr:not(.ts-sheet-total)').count()) - 1;
    await t.fillHours(last, [2]);
    await t.click('Submit');
    await page.waitForTimeout(500);
    expect(await page.isVisible('.ts-issues .is-error:has-text("Ticket #")'), 'required column error not shown');
    log('submit blocked until "Ticket #" is filled in');
    await page.locator('table.ts-sheet tbody tr:not(.ts-sheet-total)').nth(last).locator('input[aria-label="Ticket #"]').fill('APP-77');
    await t.click('Submit');
    await page.waitForSelector('.ms-Dialog-main');
    await t.dialogButton('Submit');
    await t.toast(/Submitted/);
    const v = await t.lists(() => window.__tsLists().BrillTS_TimeEntries.items.filter((x) => /APP-77/.test(x.CustomValues || '')).length);
    expect(v === 1, 'custom value not stored');
    log('value stored with the line and the week submitted');
  });

  await journey('J21', 'Holidays: a holiday is marked on the grid and lowers the weekly target', browser, async (page, t, log) => {
    await t.open('role=member', 'timesheet');
    const before = (await page.locator('.ts-kpi').filter({ hasText: 'Target' }).locator('.ts-kpi-value').textContent()).trim();
    await t.signIn('admin', 'admin/tracking');
    const field = page.locator('.ms-DatePicker').first();
    const width = (await field.boundingBox()).width;
    expect(width >= 170, `holiday date field is ${width}px wide`);
    await field.click();
    const wed = await page.evaluate(() => {
      const d = new Date();
      const monday = new Date(d.getFullYear(), d.getMonth(), d.getDate() - ((d.getDay() + 6) % 7));
      const w = new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + 2);
      const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      return { label: `${w.getDate()}, ${months[w.getMonth()]}, ${w.getFullYear()}`, key: `${w.getFullYear()}-${('0' + (w.getMonth() + 1)).slice(-2)}-${('0' + w.getDate()).slice(-2)}` };
    });
    const firstDay = (await page.locator('.ms-Callout th, .ms-Callout [role="columnheader"]').first().textContent()).trim();
    expect(/^M/.test(firstDay), 'calendar does not start on Monday: ' + firstDay);
    await page.locator(`.ms-Callout button[aria-label="${wed.label}"]`).first().click();
    expect(width >= 170, 'date field');
    await page.getByLabel('Name', { exact: true }).fill('Founders day');
    await page.locator('button:has-text("Add")').filter({ hasNotText: 'Add a' }).first().click();
    await saveAdmin(page, t);
    log(`holiday added on ${wed.key}`);
    await t.signIn('member', 'timesheet');
    expect(await page.isVisible('th.ts-day:has-text("Holiday")'), 'holiday not marked on the grid');
    const after = (await page.locator('.ts-kpi').filter({ hasText: 'Target' }).locator('.ts-kpi-value').textContent()).trim();
    expect(before === '40 h' && after === '32 h', `target ${before} → ${after}`);
    log(`grid marks the holiday; weekly target ${before} → ${after}`);
  });

  await journey('J22', 'Rule-based limits: external users are blocked above 8 h a day', browser, async (page, t, log) => {
    await t.open('role=admin', 'admin/hours');
    await t.click('Add a rule');
    const card = page.locator('.ts-card').filter({ has: page.locator('input[aria-label="Rule name"]') }).first();
    await card.locator('input[aria-label="Rule name"]').fill('Contractors');
    await card.locator('.ms-Checkbox:has-text("Only external users") label').click();
    await card.locator('.ms-Checkbox:has-text("Block submitting when broken") label').click();
    await card.getByLabel('Daily maximum').fill('8');
    await saveAdmin(page, t);
    log('rule "Contractors": external users, max 8 h a day, enforced');
    await t.signIn('ana', 'timesheet');
    expect(await page.isVisible('text=Rule: Contractors'), 'rule not shown as the source of the limits');
    await t.fillHours(0, [9, 8, 8, 8, 8]);
    await t.click('Submit');
    await page.waitForTimeout(500);
    expect(await page.isVisible('.ts-issues .is-error:has-text("daily maximum of 8 h")'), 'rule not enforced');
    log('Ana (external) sees "Rule: Contractors" and cannot submit 9 h on Monday');
    await t.signIn('member', 'timesheet');
    expect(!(await page.isVisible('text=Rule: Contractors')), 'rule applied to a member');
    log('members keep the organisation limits');
  });

  await journey('J23', 'Per-person working days: Saturday becomes a working day for one person', browser, async (page, t, log) => {
    await t.open('role=admin', 'people');
    await page.locator('td:has-text("Rahul Sharma")').click();
    await page.locator('.ms-Panel-main .ms-Checkbox:has-text("Sat") label').click();
    await page.locator('.ms-Panel-main button:has-text("Save")').click();
    await t.toast(/Saved/);
    await t.signIn('member', 'timesheet');
    const sat = await page.locator('table.ts-sheet th.ts-day').nth(5).getAttribute('class');
    const sun = await page.locator('table.ts-sheet th.ts-day').nth(6).getAttribute('class');
    expect(!/is-off/.test(sat) && /is-off/.test(sun), `Sat "${sat}" Sun "${sun}"`);
    expect(await page.isVisible('.ts-kpi-hint:has-text("6 working days")'), 'not counted as 6 working days');
    log('Rahul: Saturday is a working day (6 working days this week); Sunday stays off');
  });

  await journey('J24', 'Email templates: an edited subject is what approvers receive', browser, async (page, t, log) => {
    await t.open('role=admin', 'admin/notifications');
    const subject = page.locator('.ts-tpl-grid').getByLabel('Subject');
    await subject.fill('Please approve {{person}} ({{hours}} h)');
    await saveAdmin(page, t);
    log('“Timesheet submitted” subject changed');
    await t.signIn('member', 'timesheet');
    await t.click('Submit');
    await page.waitForSelector('.ms-Dialog-main');
    await t.dialogButton('Submit');
    await t.toast(/Submitted/);
    const titles = await t.lists(() => window.__tsLists().BrillTS_Notifications.items.filter((x) => x.Kind === 'sheetSubmitted').map((x) => x.Title));
    expect(titles.some((x) => /^Please approve Rahul Sharma \(\d+(\.\d+)? h\)$/.test(x)), 'titles: ' + titles.slice(-3).join(' | '));
    const mail = await t.lists(() => window.__tsLists().BrillTS_EmailQueue.items.filter((x) => /^Please approve Rahul/.test(x.Title)).length);
    expect(mail >= 1, 'email not queued with the new subject');
    log('notification and email use the new subject');
  });

  await journey('J25', 'Programmes: a new programme makes its manager a Programme Manager', browser, async (page, t, log) => {
    await t.open('role=admin', 'programs');
    await t.click('New programme');
    await page.locator('.ms-Panel-main').getByLabel('Programme', { exact: true }).fill('Retail');
    const picker = page.locator('.ms-Panel-main .ts-picker input');
    await picker.fill('Meera');
    await page.waitForTimeout(300);
    await picker.press('Enter');
    await page.locator('.ms-Panel-main button:has-text("Save")').click();
    await t.toast(/Programme saved/);
    const role = await t.lists(() => window.__tsLists().BrillTS_Users.items.filter((x) => x.Email === 'meera.iyer@contoso.com')[0].UserRole);
    expect(role === 'programManager', 'role ' + role);
    log('programme "Retail" created; Meera is now a Programme Manager');
    await t.signIn('manager', 'programs');
    expect(await page.isVisible('.ts-nav-item:has-text("Programmes")') && (await page.isVisible('text=Retail')), 'Meera does not see her programme');
    log('Meera sees Programmes and her programme');
  });

  await journey('J26', 'People: add a person; the Standard user limit warns when exceeded', browser, async (page, t, log) => {
    await t.open('role=admin&plan=standard', 'people');
    for (const [name, email] of [['Nina Patel', 'nina.patel@contoso.com'], ['Omar Farouk', 'omar.farouk@contoso.com'], ['Lea Martin', 'lea.martin@contoso.com']]) {
      await t.click('Add person');
      await page.locator('.ms-Panel-main').getByLabel('Name', { exact: true }).fill(name);
      await page.locator('.ms-Panel-main').getByLabel('Email', { exact: true }).fill(email);
      await page.locator('.ms-Panel-main button:has-text("Save")').click();
      await t.toast(/Saved/);
      await page.waitForTimeout(400);
    }
    const n = await t.lists(() => window.__tsLists().BrillTS_Users.items.filter((x) => x.IsActive !== false).length);
    expect(n === 11, 'active users ' + n);
    log('3 people added (11 active)');
    expect(await page.isVisible('text=beyond the 10 users'), 'user limit warning not shown');
    log('Standard plan (10 users): warning that 1 person is beyond the limit');
  });

  await journey('J27', 'Approvals: send back several weeks at once', browser, async (page, t, log) => {
    await t.open('role=manager', 'approvals');
    const n = await page.locator('table.ts-table tbody tr').count();
    expect(n >= 2, 'need at least 2 waiting, have ' + n);
    await page.locator('thead [aria-label="Select all"]').click({ force: true });
    await page.locator('.ts-bulkbar button:has-text("Send back")').click();
    await page.locator('.ms-Dialog-main textarea').fill('Please add task details');
    await t.dialogButton('Send back');
    await t.toast(/sent back/);
    const rejected = await t.lists(() => window.__tsLists().BrillTS_Timesheets.items.filter((x) => x.SheetStatus === 'Rejected' && x.DecidedBy === 'Meera Iyer').length);
    expect(rejected === n, `${rejected} of ${n} sent back`);
    const notes = await t.lists(() => window.__tsLists().BrillTS_Notifications.items.filter((x) => x.Kind === 'sheetRejected').length);
    expect(notes === n, 'people not told');
    log(`${n} weeks sent back with one reason; each person notified`);
  });

  await journey('J28', 'My timesheets: statuses, filter and CSV export', browser, async (page, t, log) => {
    await t.open('role=member', 'history');
    const approved = Number((await page.locator('.ts-kpi').filter({ hasText: 'Approved' }).locator('.ts-kpi-value').textContent()).trim());
    expect(approved >= 2, 'approved ' + approved);
    await page.locator('.ts-kpi').filter({ hasText: 'Approved' }).click();
    await page.waitForTimeout(300);
    const rows = await page.locator('table.ts-table tbody tr').count();
    expect(rows === approved, `filter shows ${rows}, expected ${approved}`);
    log(`${approved} approved weeks; the Approved filter shows exactly those`);
    await page.locator('[aria-label="More export formats"]').first().click();
    const [d] = await Promise.all([page.waitForEvent('download'), page.locator('button[role="menuitem"]:has-text("CSV")').click()]);
    const file = path.join(OUT, d.suggestedFilename());
    await d.saveAs(file);
    const csv = fs.readFileSync(file, 'utf8');
    expect(/Week starting,Status,Total hours/.test(csv) && csv.split('\n').length > 3, 'csv content');
    log(`CSV export: ${d.suggestedFilename()} (${csv.split('\n').length - 1} weeks)`);
  });

  await journey('J29', 'Dashboards: add a widget and keep it after saving', browser, async (page, t, log) => {
    await t.open('role=admin', 'dashboards');
    await t.click('Save as my dashboard');
    await page.locator('.ms-Panel-main').getByLabel('Name').fill('Delivery overview');
    await page.locator('.ms-Panel-main button:has-text("Add a widget")').click();
    const titles = page.locator('.ms-Panel-main').getByLabel('Title');
    await titles.last().fill('Hours by department');
    await page.locator('.ms-Panel-main button:has-text("Save dashboard")').click();
    await t.toast(/Dashboard saved/);
    await page.waitForFunction(() => document.querySelectorAll('.ts-widgets .ts-card').length === 7, undefined, { timeout: 10000 }).catch(() => undefined);
    const saved = await t.lists(() => window.__tsLists().BrillTS_SavedReports.items.filter((x) => x.Title === 'Delivery overview').map((x) => JSON.parse(x.Definition).widgets.length));
    expect(saved.length === 1 && saved[0] === 7, 'saved widgets ' + saved);
    expect((await page.locator('.ts-widgets .ts-card').count()) === 7 && (await page.isVisible('text=Hours by department')), 'widget not shown');
    log('dashboard "Delivery overview" saved with 7 widgets, the new one shown');
  });

  await journey('J30', 'Report filters: billable only, one person', browser, async (page, t, log) => {
    await t.open('role=admin', 'reports');
    await t.dropdownByLabel('Period', 'Last 90 days');
    await page.waitForTimeout(800);
    const kpi = async (label) => Number((await page.locator('.ts-kpi').filter({ has: page.locator('.ts-kpi-label', { hasText: new RegExp(`^${label}$`) }) }).first().locator('.ts-kpi-value').textContent()).replace(/[^\d.]/g, ''));
    const all = await kpi('Hours');
    await t.dropdownByLabel('Billable', 'Billable only');
    await page.waitForTimeout(500);
    expect((await kpi('Non-billable')) === 0 && (await kpi('Hours')) < all, 'billable filter');
    log(`billable only: ${await kpi('Hours')} of ${all} h, non-billable 0`);
    await t.dropdownByLabel('Billable', 'All hours');
    await page.locator('.ms-Dropdown-container').filter({ has: page.locator('label', { hasText: /^People$/ }) }).locator('.ms-Dropdown').click();
    await page.locator('.ms-Checkbox:has-text("Sara Khan")').first().click();
    await page.keyboard.press('Escape');
    await page.waitForTimeout(800);
    const sara = await t.lists(() => window.__tsLists().BrillTS_TimeEntries.items.filter((x) => x.UserEmail === 'sara.khan@contoso.com').reduce((s, e) => s + Object.values(JSON.parse(e.DayHours || '{}')).reduce((a, b) => a + b, 0), 0));
    const shown = await kpi('Hours');
    expect(Math.abs(shown - sara) < 0.01, `Sara ${shown} h shown vs ${sara} h in the data`);
    log(`one person: ${shown} h, matching Sara's data`);
  });

  await journey('J31', 'Task Manager: reopening an approved week takes the hours back out', browser, async (page, t, log) => {
    await t.open('role=admin&tm=1', 'projects');
    await t.toast(/Task Manager:/, 15000);
    await t.signIn('member', 'timesheet');
    await t.click('Add line');
    await t.pick('Project', 'Customer portal');
    await t.pick('Task', 'Sign-in and accounts');
    const last = (await page.locator('table.ts-sheet tbody tr:not(.ts-sheet-total)').count()) - 1;
    await t.fillHours(last, [3, 3]);
    await t.click('Submit');
    await t.dialogButton('Submit');
    await t.toast(/Submitted/);
    const spent = () => t.lists(() => window.__tsLists().TM_Tasks.items.filter((x) => x.TaskNumber === 'TM-101')[0].SpentMinutes);
    const s0 = await spent();
    await t.signIn('manager', 'approvals');
    await page.locator('table.ts-table tbody tr').filter({ hasText: 'Rahul Sharma' }).filter({ hasText: 'just now' }).first().click();
    await page.locator('.ms-Panel-main button:has-text("Approve")').click();
    await t.toast(/approved/);
    await page.waitForTimeout(800);
    const s1 = await spent();
    expect(s1 - s0 === 360, `approve ${s0} → ${s1}`);
    await page.locator('button[role="tab"]:has-text("Decided")').click();
    await page.waitForTimeout(1200);
    await page.locator('table.ts-table tbody tr').filter({ hasText: 'Rahul Sharma' }).filter({ hasText: 'Approved' }).first().click();
    await page.locator('.ms-Panel-main button:has-text("Reopen")').click();
    await page.locator('.ms-Dialog-main textarea').fill('Wrong task');
    await t.dialogButton('Reopen');
    await t.toast(/Reopened/);
    await page.waitForTimeout(800);
    const s2 = await spent();
    expect(s2 === s0, `reopen ${s1} → ${s2}`);
    log(`approved: +360 min in Task Manager; reopened: back to ${s2} min`);
  });

  await journey('J32', 'Housekeeping: delete a task and an activity, close a support ticket', browser, async (page, t, log) => {
    await t.open('role=admin', 'projects');
    await page.locator('td:has-text("Website redesign")').first().click();
    await page.locator('button[role="tab"]:has-text("Tasks")').click();
    const before = await page.locator('table.ts-table tbody tr').count();
    await page.locator('tr:has-text("Accessibility review") button[aria-label="Delete task"]').click();
    await t.dialogButton('Delete');
    await t.toast(/Task deleted/);
    expect((await page.locator('table.ts-table tbody tr').count()) === before - 1, 'task still listed');
    log('task deleted');
    await t.go('activities');
    await page.locator('tr:has-text("Training") button:has-text("Delete")').click();
    await t.dialogButton('Delete');
    await t.toast(/Deleted/);
    expect(!(await t.lists(() => window.__tsLists().BrillTS_Activities.items.some((x) => x.Title === 'Training'))), 'activity not deleted');
    log('activity deleted');
    await t.go('support');
    await page.getByLabel('Subject').fill('How do I export to Power BI?');
    await page.getByLabel('Details').fill('Details');
    await t.click('Send ticket');
    await t.toast(/Ticket sent/);
    await page.locator('tr:has-text("How do I export to Power BI?") button:has-text("Close")').click();
    await page.waitForTimeout(800);
    expect(await page.isVisible('tr:has-text("How do I export to Power BI?") .ts-status:has-text("Closed")'), 'ticket not closed');
    log('support ticket raised and closed');
  });

  await journey('J33', 'A project manager’s own week goes to their manager, never to themselves', browser, async (page, t, log) => {
    await t.open('role=pm', 'timesheet');
    await t.pick('Project', 'Website redesign');
    await t.fillHours(0, [8, 8, 8, 8, 8]);
    await t.click('Submit');
    await t.dialogButton('Submit');
    await t.toast(/Submitted for approval/);
    const to = await t.lists(() => window.__tsLists().BrillTS_Notifications.items.filter((x) => x.Kind === 'sheetSubmitted' && /Arjun/.test(x.Title)).map((x) => x.Recipient));
    expect(to.indexOf('vikram.rao@contoso.com') >= 0 && to.indexOf('arjun.mehta@contoso.com') < 0, 'approvers: ' + to);
    log('Arjun’s week went to Vikram (his manager), not to Arjun as project manager');
    await t.signIn('pm', 'approvals');
    const own = await page.locator('table.ts-table tbody tr').filter({ hasText: 'Arjun Mehta' }).count();
    expect(own === 0, 'Arjun can approve his own week');
    log('Arjun does not see his own week to approve');
  });

  await browser.close();
  const passed = results.filter((r) => r.ok).length;
  console.log(`\n${passed} of ${results.length} journeys passed, ${results.reduce((s, r) => s + r.steps.length, 0)} checks`);
  fs.writeFileSync(path.join(OUT, 'results.json'), JSON.stringify(results, null, 2));
  process.exit(passed === results.length ? 0 : 1);
}

main().catch((e) => {
  console.error(`aborted in ${current}:`, e);
  process.exit(1);
});

// Builds the client product guide from docs/client/guide.html:
//   docs/client/dist/Brillienta-Timesheet-Product-Guide.html  self-contained (screenshots embedded), for sending
//   docs/client/dist/Brillienta-Timesheet-Product-Guide.pdf   A4, printed with headless Chrome
// Usage: node tools/docs/build-guide.cjs [--fragment <file>]   (--fragment also writes the page body without
// <html>/<head> wrappers, for publishing as a hosted page)
// Screenshots: npm run preview, then node tools/docs/capture.cjs
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const root = path.resolve(__dirname, '../..');
const srcDir = path.join(root, 'docs/client');
const outDir = path.join(srcDir, 'dist');
const name = 'Brillienta-Timesheet-Product-Guide';

const source = fs.readFileSync(path.join(srcDir, 'guide.html'), 'utf8');
let missing = 0;
const body = source.replace(/src="(shots\/[^"]+\.png)"/g, (_, rel) => {
  const file = path.join(srcDir, rel);
  if (!fs.existsSync(file)) {
    console.error('Missing screenshot: ' + rel);
    missing++;
    return `src="${rel}"`;
  }
  return `src="data:image/png;base64,${fs.readFileSync(file).toString('base64')}"`;
});
if (missing) process.exit(1);

fs.mkdirSync(outDir, { recursive: true });
const page =
  '<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n' +
  body.replace(/<\/style>/, '</style>\n</head>\n<body>') +
  '\n</body>\n</html>\n';
const htmlOut = path.join(outDir, name + '.html');
fs.writeFileSync(htmlOut, page);
console.log(`HTML: ${path.relative(root, htmlOut)} (${(page.length / 1048576).toFixed(1)} MB)`);

const fragArg = process.argv.indexOf('--fragment');
if (fragArg > 0 && process.argv[fragArg + 1]) {
  fs.writeFileSync(process.argv[fragArg + 1], body);
  console.log('Fragment: ' + process.argv[fragArg + 1]);
}

const pdfTarget = path.join(outDir, name + '.pdf');
let playwright;
try {
  playwright = require('playwright');
} catch {
  playwright = undefined;
}
if (playwright && !process.env.CHROME) {
  // Playwright's bundled browser prints reliably everywhere (no Chrome path, no first-run profile quirks).
  (async () => {
    const browser = await playwright.chromium.launch();
    const page = await browser.newPage();
    await page.goto('file://' + htmlOut.split(path.sep).join('/'), { waitUntil: 'load' });
    await page.waitForTimeout(1500);
    await page.pdf({ path: pdfTarget, preferCSSPageSize: true, printBackground: true });
    await browser.close();
    console.log(`PDF:  ${path.relative(root, pdfTarget)} (${(fs.statSync(pdfTarget).size / 1048576).toFixed(1)} MB)`);
  })().catch((e) => {
    console.error('PDF failed: ' + e.message);
    process.exit(1);
  });
  return;
}
const chromes = [
  process.env.CHROME,
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
  '/usr/bin/google-chrome',
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
].filter(Boolean);
const chrome = chromes.find((c) => fs.existsSync(c));
if (!chrome) {
  console.log('No Chrome or Edge found (set CHROME=...): PDF skipped.');
  process.exit(0);
}
const pdfOut = path.join(outDir, name + '.pdf');
// A reused profile: headless Chrome often skips printing on the very first run of a brand-new profile.
const profile = path.join(require('os').tmpdir(), 'brillienta-timesheet-guide-pdf');
const url = 'file:///' + htmlOut.split(path.sep).join('/');
let log = '';
if (fs.existsSync(pdfOut)) fs.unlinkSync(pdfOut);
for (let attempt = 1; attempt <= 3 && !fs.existsSync(pdfOut); attempt++) {
  const run = spawnSync(
    chrome,
    ['--headless=new', '--disable-gpu', '--no-pdf-header-footer', '--virtual-time-budget=15000', `--user-data-dir=${profile}`, `--print-to-pdf=${pdfOut}`, url],
    { encoding: 'utf8', timeout: 180000 }
  );
  log = String(run.stderr || run.error || '');
}
if (!fs.existsSync(pdfOut)) {
  console.error('Chrome did not write the PDF.\n' + log.slice(-2000));
  process.exit(1);
}
console.log(`PDF:  ${path.relative(root, pdfOut)} (${(fs.statSync(pdfOut).size / 1048576).toFixed(1)} MB)`);

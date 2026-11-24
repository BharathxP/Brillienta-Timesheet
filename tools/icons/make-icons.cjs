#!/usr/bin/env node
// Renders the clock logo (assets/defaultLogo.ts) to the PNGs the package needs, with the preinstalled Chromium:
//   sharepoint/images/app-icon.png (96×96, app catalog) and teams/<web part id>_color.png / _outline.png.
// With --brand: docs/brand/ (logo SVG, 512 px PNG, badge SVG) only.
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');

const WEB_PART_ID = 'ff5b1593-a3b3-45f5-b3a5-8d11821fc2a1';
const svg = (bg, face, hands, accent) =>
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120"><rect width="120" height="120" rx="28" fill="${bg}"/>` +
  `<circle cx="60" cy="62" r="34" fill="${face}"/><rect x="54" y="16" width="12" height="10" rx="3" fill="${accent}"/>` +
  `<path d="M60 62 V40" stroke="${hands}" stroke-width="7" stroke-linecap="round"/><path d="M60 62 L76 72" stroke="${hands}" stroke-width="7" stroke-linecap="round"/>` +
  `<circle cx="60" cy="62" r="6" fill="${accent}" stroke="${hands}" stroke-width="3"/></svg>`;
const outline =
  `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 120 120"><circle cx="60" cy="62" r="34" fill="none" stroke="#fff" stroke-width="8"/>` +
  `<path d="M60 62 V40 M60 62 L76 72" stroke="#fff" stroke-width="8" stroke-linecap="round"/><rect x="54" y="14" width="12" height="10" rx="3" fill="#fff"/></svg>`;

(async () => {
  const root = path.join(__dirname, '..', '..');
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const shot = async (markup, size, file, transparent) => {
    await page.setViewportSize({ width: size, height: size });
    await page.setContent(`<html><body style="margin:0;background:transparent">${markup.replace('<svg ', `<svg width="${size}" height="${size}" `)}</body></html>`);
    await page.screenshot({ path: path.join(root, file), omitBackground: !!transparent });
    console.log('wrote', file);
  };
  if (process.argv.includes('--brand')) {
    // docs/brand: the logo for the website, listings and documents (the package icons are left as they are).
    const dir = path.join(root, 'docs/brand');
    fs.mkdirSync(dir, { recursive: true });
    const logo = svg('#038f62', '#ffffff', '#17232e', '#caf578');
    fs.writeFileSync(path.join(dir, 'brillienta-timesheet-logo.svg'), logo + '\n');
    // The clock alone (navy hands, green crown), for places where a tile does not fit.
    const badge = logo.replace(/<rect width="120" height="120" rx="28" fill="#038f62"\/>/, '').replace('fill="#ffffff"', 'fill="#ffffff" stroke="#038f62" stroke-width="6"').replace(/fill="#caf578"/g, 'fill="#038f62"');
    fs.writeFileSync(path.join(dir, 'brillienta-timesheet-badge.svg'), badge + '\n');
    await shot(logo, 512, 'docs/brand/brillienta-timesheet-logo-512.png', true);
    await browser.close();
    return;
  }
  fs.mkdirSync(path.join(root, 'sharepoint/images'), { recursive: true });
  fs.mkdirSync(path.join(root, 'teams'), { recursive: true });
  await shot(svg('#038f62', '#ffffff', '#17232e', '#caf578'), 96, 'sharepoint/images/app-icon.png', true);
  await shot(svg('#038f62', '#ffffff', '#17232e', '#caf578'), 192, `teams/${WEB_PART_ID}_color.png`, true);
  await shot(outline, 32, `teams/${WEB_PART_ID}_outline.png`, true);
  await browser.close();
})();

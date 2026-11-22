#!/usr/bin/env node
// Renders the clock logo (assets/defaultLogo.ts) to the PNGs the package needs, with the preinstalled Chromium:
//   sharepoint/images/app-icon.png (96×96, app catalog) and teams/<web part id>_color.png / _outline.png.
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
  fs.mkdirSync(path.join(root, 'sharepoint/images'), { recursive: true });
  fs.mkdirSync(path.join(root, 'teams'), { recursive: true });
  await shot(svg('#038f62', '#ffffff', '#17232e', '#caf578'), 96, 'sharepoint/images/app-icon.png', true);
  await shot(svg('#038f62', '#ffffff', '#17232e', '#caf578'), 192, `teams/${WEB_PART_ID}_color.png`, true);
  await shot(outline, 32, `teams/${WEB_PART_ID}_outline.png`, true);
  await browser.close();
})();

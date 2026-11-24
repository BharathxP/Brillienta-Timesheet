// The Microsoft Teams app manifest shipped in the package as teams/manifest.json, next to the icons
// teams/<web part id>_color.png and _outline.png (the layout Brillienta-Helpdesk and Employee Directory use). The packager
// copies teams/ into the .sppkg, and Sync to Teams deploys this manifest instead of generating one from the web part (a
// generated one is named after the web part title and lacks supportsChannelFeatures, which Teams requires).
//
// The manifest itself is built by the app's own code (services/packages.ts › teamsManifest, from services/product.ts),
// the same function behind Admin › Integrations › Teams › Download, so both are always identical.
//
// buildManifest()  → the manifest object
// buildPackage()   → a zip with manifest.json and the icons at its root, for Teams admin center
// checkPackage()   → problems against the official Teams schema and the app package rules (empty when valid)
//
// Written by `node tools/teams/make-app-package.cjs` (after `npm run build:dev`); checked by tests/teams.test.cjs.
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..', '..');
const read = (p) => fs.readFileSync(path.join(root, p));
const json = (p) => JSON.parse(read(p).toString('utf8'));
const lib = (p) => require(path.join(root, 'lib-commonjs/webparts/brillientaTimesheet', p));

const { TEAMS_MANIFEST_VERSION: MANIFEST_VERSION, SHAREPOINT_APP_ID } = lib('services/product.js');

function iconNames(id) {
  return lib('services/packages.js').teamsIconNames(id);
}

function sources() {
  const webPart = json('src/webparts/brillientaTimesheet/BrillientaTimesheetWebPart.manifest.json');
  return {
    webPart,
    version: json('package.json').version,
    color: read(`teams/${iconNames(webPart.id).color}`),
    outline: read(`teams/${iconNames(webPart.id).outline}`)
  };
}

function buildManifest(s = sources()) {
  return lib('services/packages.js').teamsManifest(s.version);
}

/** Fixed date, so the zip only changes when its content does. */
const ZIP_DATE = new Date(2026, 0, 1, 0, 0, 0);

function buildPackage(s = sources()) {
  const { zipFiles } = lib('services/zip.js');
  return Buffer.from(
    zipFiles(
      [
        { path: 'manifest.json', content: manifestText(s) },
        { path: iconNames(s.webPart.id).color, content: new Uint8Array(s.color) },
        { path: iconNames(s.webPart.id).outline, content: new Uint8Array(s.outline) }
      ],
      ZIP_DATE
    )
  );
}

/** Reads a PNG's size, colour type and decoded RGBA pixels (8-bit, non-interlaced: what our icons are). */
function png(buf) {
  const zlib = require('zlib');
  if (buf.toString('ascii', 1, 4) !== 'PNG') throw new Error('not a PNG');
  let at = 8;
  let w = 0;
  let h = 0;
  let type = 0;
  const idat = [];
  while (at < buf.length) {
    const len = buf.readUInt32BE(at);
    const kind = buf.toString('ascii', at + 4, at + 8);
    const data = buf.subarray(at + 8, at + 8 + len);
    if (kind === 'IHDR') {
      w = data.readUInt32BE(0);
      h = data.readUInt32BE(4);
      type = data[9];
      if (data[8] !== 8 || data[12] !== 0) throw new Error('only 8-bit non-interlaced PNGs are read');
    } else if (kind === 'IDAT') idat.push(data);
    at += 12 + len;
  }
  const ch = { 2: 3, 6: 4 }[type];
  if (!ch) throw new Error(`PNG colour type ${type} not read`);
  const raw = zlib.inflateSync(Buffer.concat(idat));
  const stride = w * ch;
  const px = Buffer.alloc(w * h * 4);
  let prev = Buffer.alloc(stride);
  for (let y = 0; y < h; y++) {
    const f = raw[y * (stride + 1)];
    const line = Buffer.from(raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1)));
    for (let i = 0; i < stride; i++) {
      const a = i >= ch ? line[i - ch] : 0;
      const b = prev[i];
      const c = i >= ch ? prev[i - ch] : 0;
      const p = a + b - c;
      const pa = Math.abs(p - a);
      const pb = Math.abs(p - b);
      const pc = Math.abs(p - c);
      const paeth = pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
      line[i] = (line[i] + [0, a, b, (a + b) >> 1, paeth][f]) & 0xff;
    }
    for (let x = 0; x < w; x++) {
      const o = (y * w + x) * 4;
      px[o] = line[x * ch];
      px[o + 1] = line[x * ch + 1];
      px[o + 2] = line[x * ch + 2];
      px[o + 3] = ch === 4 ? line[x * ch + 3] : 255;
    }
    prev = line;
  }
  return { w, h, type, px };
}

/**
 * Problems with the manifest and icons against the Teams rules: the official JSON schema for the manifest version
 * (from @microsoft/app-manifest) plus the app package rules the schema doesn't express. Empty when valid.
 */
function checkPackage(manifest, color, outline) {
  const problems = [];
  const Ajv = require('ajv-draft-04');
  const addFormats = require('ajv-formats');
  const schema = require(`@microsoft/app-manifest/build/json-schemas/teams/v${MANIFEST_VERSION}/MicrosoftTeams.schema.json`);
  const ajv = new Ajv({ allErrors: true, strict: false });
  addFormats(ajv);
  const validate = ajv.compile(schema);
  if (!validate(manifest)) for (const e of validate.errors) problems.push(`schema: ${e.instancePath || '/'} ${e.message}${e.params && e.params.additionalProperty ? ` (${e.params.additionalProperty})` : ''}`);

  if (!/^\d+\.\d+\.\d+$/.test(manifest.version)) problems.push('version must be major.minor.patch');
  if (manifest.name.short.length > 30) problems.push('name.short over 30 characters');
  if (manifest.description.full === manifest.description.short) problems.push('description.full repeats description.short');
  for (const d of manifest.validDomains) if (d === '*' || /^\*\.[^.]+$/.test(d)) problems.push(`validDomains: ${d} is too broad`);
  if ((manifest.configurableTabs || []).length && manifest.supportsChannelFeatures !== 'tier1') problems.push('channel tabs need supportsChannelFeatures: tier1 (Teams rejects the upload without it)');
  if (manifest.supportedChannelTypes) problems.push('supportedChannelTypes is not accepted in manifest v1.25+, use supportsChannelFeatures');
  if (!manifest.webApplicationInfo || manifest.webApplicationInfo.id !== SHAREPOINT_APP_ID) problems.push('webApplicationInfo.id must be SharePoint Online');

  const c = png(color);
  if (c.w !== 192 || c.h !== 192) problems.push(`colour icon is ${c.w}×${c.h}, must be 192×192`);
  const o = png(outline);
  if (o.w !== 32 || o.h !== 32) problems.push(`outline icon is ${o.w}×${o.h}, must be 32×32`);
  if (o.type !== 6) problems.push('outline icon needs an alpha channel');
  let opaque = 0;
  for (let i = 0; i < o.px.length; i += 4) {
    const a = o.px[i + 3];
    if (a === 0) continue;
    opaque++;
    if (o.px[i] < 250 || o.px[i + 1] < 250 || o.px[i + 2] < 250) {
      problems.push('outline icon must be white on transparent');
      break;
    }
  }
  if (!opaque) problems.push('outline icon is empty');
  return problems;
}

/** Problems with a Copilot declarative agent file against the official schema (copilot/declarative-agent/v1.2). */
function checkDeclarativeAgent(agent) {
  const Ajv = require('ajv-draft-04');
  const addFormats = require('ajv-formats');
  const schema = require('@microsoft/app-manifest/build/json-schemas/copilot/declarative-agent/v1.2/schema.json');
  const ajv = new Ajv({ allErrors: true, strict: false });
  addFormats(ajv);
  const validate = ajv.compile(schema);
  return validate(agent) ? [] : validate.errors.map((e) => `agent schema: ${e.instancePath || '/'} ${e.message}${e.params && e.params.additionalProperty ? ` (${e.params.additionalProperty})` : ''}`);
}

/** Problems with the Teams manifest alone (schema and rules), without icons: for the Copilot agent's app. */
function checkManifestOnly(manifest) {
  const Ajv = require('ajv-draft-04');
  const addFormats = require('ajv-formats');
  const schema = require(`@microsoft/app-manifest/build/json-schemas/teams/v${MANIFEST_VERSION}/MicrosoftTeams.schema.json`);
  const ajv = new Ajv({ allErrors: true, strict: false });
  addFormats(ajv);
  const validate = ajv.compile(schema);
  return validate(manifest) ? [] : validate.errors.map((e) => `schema: ${e.instancePath || '/'} ${e.message}${e.params && e.params.additionalProperty ? ` (${e.params.additionalProperty})` : ''}`);
}

/** The manifest as written to teams/manifest.json. */
function manifestText(s = sources()) {
  return JSON.stringify(buildManifest(s), null, 2) + '\n';
}

module.exports = { buildManifest, buildPackage, checkPackage, checkManifestOnly, checkDeclarativeAgent, manifestText, iconNames, sources, png, MANIFEST_VERSION, SHAREPOINT_APP_ID };

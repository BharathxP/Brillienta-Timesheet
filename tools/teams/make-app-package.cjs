// Writes teams/manifest.json, the Teams app manifest that goes into the .sppkg with the icons in teams/ (see
// app-manifest.cjs), after checking it against the Teams manifest rules. Run after `npm run build:dev` and whenever the
// version, name, descriptions, developer details or icons change:
//   node tools/teams/make-app-package.cjs
// With --zip <file> it also writes a Teams app package (manifest and icons at the zip root) for uploading in Teams
// admin center › Manage apps › Upload new app:
//   node tools/teams/make-app-package.cjs --zip release/teams-app-package.zip
const fs = require('fs');
const path = require('path');
const { buildManifest, buildPackage, checkPackage, manifestText, sources } = require('./app-manifest.cjs');

const s = sources();
const manifest = buildManifest(s);
const problems = checkPackage(manifest, s.color, s.outline);
if (problems.length) {
  console.error('Teams app manifest is not valid:\n- ' + problems.join('\n- '));
  process.exit(1);
}
const root = path.resolve(__dirname, '..', '..');
fs.writeFileSync(path.join(root, 'teams', 'manifest.json'), manifestText(s));
console.log(`wrote teams/manifest.json (Teams app ${manifest.name.short} ${manifest.version}), valid`);

const z = process.argv.indexOf('--zip');
if (z > 0 && process.argv[z + 1]) {
  const out = path.resolve(process.argv[z + 1]);
  fs.mkdirSync(path.dirname(out), { recursive: true });
  fs.writeFileSync(out, buildPackage(s));
  console.log(`wrote ${path.relative(root, out)}`);
}

// Builds the local preview (npm run preview). Uses webpack's Node API, so no extra CLI package is needed.
const fs = require('fs');
const path = require('path');
const webpack = require('webpack');
const config = require('./webpack.config.cjs');

const out = path.resolve(__dirname, '../preview-dist');
fs.mkdirSync(out, { recursive: true });
['index.html'].forEach((f) => fs.copyFileSync(path.resolve(__dirname, f), path.join(out, f)));

webpack(config, (err, stats) => {
  if (err || stats.hasErrors()) {
    console.error(err || stats.toString({ colors: false, all: false, errors: true }));
    process.exit(1);
  }
  console.log(`Preview built: ${path.join(out, 'index.html')}`);
  console.log('Open it in a browser (file://). Add ?role=member, ?role=manager or ?role=pm to see other roles; ?plan=standard|pro to see a plan.');
});

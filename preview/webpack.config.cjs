// Local preview bundle (npm run preview): the compiled app from lib/ with SharePoint replaced by an in-memory fake.
const path = require('path');
const webpack = require('webpack');

module.exports = {
  mode: 'development',
  devtool: false,
  entry: path.resolve(__dirname, 'entry.js'),
  output: { path: path.resolve(__dirname, '../preview-dist'), filename: 'bundle.js' },
  resolve: {
    alias: {
      '@microsoft/sp-http$': path.resolve(__dirname, 'spHttpStub.js'),
      './Timesheet.module.scss$': path.resolve(__dirname, 'styles.js')
    }
  },
  module: { rules: [{ test: /\.css$/, type: 'asset/source' }] },
  // The preview signs its own licence keys (?plan=), so it verifies them with its own public key.
  plugins: [new webpack.NormalModuleReplacementPlugin(/licensing[\\/]publicKey(\.js)?$/, path.resolve(__dirname, 'previewKey.js'))],
  performance: { hints: false }
};

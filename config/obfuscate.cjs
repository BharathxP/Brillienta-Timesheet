// Webpack patch (config/webpack-patch.json): obfuscates the licensing code in production builds so the licence checks
// are hard to find and change in the shipped bundle. Development builds are left readable.
//
// Covered: services/licensing/* (plans, licence, trial, the engine's sentinel check) and components/licensing/* (the
// screens' independent check). Not "selfDefending": the bundle is minified again afterwards, which that option breaks.

const WebpackObfuscator = require('webpack-obfuscator');

/** Compiled licensing modules (Windows or POSIX paths). */
function isLicensing(resourcePath) {
  const p = String(resourcePath || '').split('\\').join('/');
  return /\/lib\/webparts\/brillientaTimesheet\/(services|components)\/licensing\/[^/]+\.js$/.test(p);
}

module.exports = function patch(config) {
  if (config.mode !== 'production') return config;
  config.module = config.module || {};
  config.module.rules = config.module.rules || [];
  config.module.rules.push({
    test: isLicensing,
    enforce: 'post',
    use: {
      loader: WebpackObfuscator.loader,
      options: {
        compact: true,
        controlFlowFlattening: true,
        controlFlowFlatteningThreshold: 0.75,
        deadCodeInjection: true,
        deadCodeInjectionThreshold: 0.3,
        identifierNamesGenerator: 'hexadecimal',
        numbersToExpressions: true,
        renameGlobals: false,
        selfDefending: false,
        simplify: true,
        splitStrings: true,
        splitStringsChunkLength: 5,
        stringArray: true,
        stringArrayCallsTransform: true,
        stringArrayEncoding: ['rc4'],
        stringArrayIndexShift: true,
        stringArrayRotate: true,
        stringArrayShuffle: true,
        stringArrayThreshold: 1,
        stringArrayWrappersCount: 2,
        stringArrayWrappersType: 'function',
        transformObjectKeys: true,
        unicodeEscapeSequence: false,
        sourceMap: false
      }
    }
  });
  return config;
};

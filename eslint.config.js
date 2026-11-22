const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');

module.exports = [
  ...spfxProfile,
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      }
    },
    rules: {
      '@typescript-eslint/no-use-before-define': ['error', { functions: false, classes: false, variables: false, enums: false }],
      // `void promise` is how no-floating-promises wants fire-and-forget calls in event handlers marked.
      'no-void': 'off',
      // SharePoint REST clears a column only when it is sent as null, so null is part of the data contract here.
      '@rushstack/no-new-null': 'off'
    }
  }
];

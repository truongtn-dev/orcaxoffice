const js = require('@eslint/js');
const globals = require('globals');

module.exports = [
  { ignores: ['node_modules/**', 'dist/**'] },
  js.configs.recommended,
  {
    files: ['**/*.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'commonjs',
      globals: {
        ...globals.browser,
        ...globals.node,
        Office: 'readonly',
        Word: 'readonly',
        Excel: 'readonly',
      },
    },
  },
];

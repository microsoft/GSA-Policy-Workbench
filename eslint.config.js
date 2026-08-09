import js from '@eslint/js';
import globals from 'globals';
import reactHooks from 'eslint-plugin-react-hooks';
import reactRefresh from 'eslint-plugin-react-refresh';
import tseslint from 'typescript-eslint';

export default tseslint.config(
  { ignores: ['dist'] },
  {
    extends: [js.configs.recommended, ...tseslint.configs.recommended],
    files: ['**/*.{ts,tsx}'],
    languageOptions: {
      ecmaVersion: 2022,
      globals: globals.browser,
    },
    plugins: {
      'react-hooks': reactHooks,
      'react-refresh': reactRefresh,
    },
    rules: {
      ...reactHooks.configs.recommended.rules,
      'react-refresh/only-export-components': ['warn', { allowConstantExport: true }],
      // The Graph adapter is sacred: no raw fetch to graph.microsoft.com from anywhere but the adapter.
      'no-restricted-imports': 'off',
    },
  },
  {
    // Allow direct GraphClient usage only inside the adapter.
    files: ['src/**/*.{ts,tsx}'],
    ignores: ['src/adapters/graph/**'],
    rules: {
      'no-restricted-syntax': [
        'warn',
        {
          selector:
            "CallExpression[callee.object.name='Client'][callee.property.name='init']",
          message:
            'Do not construct a Graph Client outside src/adapters/graph/. Use the adapter.',
        },
      ],
    },
  },
);

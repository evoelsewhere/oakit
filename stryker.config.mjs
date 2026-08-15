import { mutatedFiles } from './scripts/mutation-scope.mjs';

/** @type {import('@stryker-mutator/api/core').PartialStrykerOptions} */
const config = {
  mutate: mutatedFiles,
  plugins: [
    '@stryker-mutator/typescript-checker',
    '@stryker-mutator/vitest-runner',
  ],
  testRunner: 'vitest',
  checkers: ['typescript'],
  coverageAnalysis: 'perTest',
  ignoreStatic: false,
  concurrency: 2,
  timeoutMS: 10_000,
  dryRunTimeoutMinutes: 3,
  incremental: true,
  reporters: ['clear-text', 'progress', 'html', 'json'],
  htmlReporter: {
    fileName: 'reports/mutation/mutation.html',
  },
  jsonReporter: {
    fileName: 'reports/mutation/mutation.json',
  },
  thresholds: {
    high: 100,
    low: 100,
    break: 100,
  },
};

export default config;

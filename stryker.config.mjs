import { mutatedFiles } from './scripts/mutation-scope.mjs';
import { focusedMutationTimeoutMs } from './scripts/mutation-timeouts.mjs';

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
  timeoutMS: focusedMutationTimeoutMs,
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
  vitest: {
    configFile: 'vitest.stryker.config.ts',
  },
};

export default config;

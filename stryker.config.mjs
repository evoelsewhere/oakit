/** @type {import('@stryker-mutator/api/core').PartialStrykerOptions} */
const config = {
  mutate: [
    'src/common/archive/read-entry.ts',
    'src/common/opc/part-uri.ts',
    'src/common/text/css.ts',
    'src/common/text/html.ts',
    'src/common/xml/normalize.ts',
    'src/common/xml/read-xml.ts',
    'src/common/xml/types.ts',
    'src/common/xml/validate.ts',
    'src/formats/pptx/internal/resource-limits.ts',
  ],
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

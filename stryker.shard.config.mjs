import process from 'node:process';

import baseConfig from './stryker.config.mjs';

const mutate = (process.env.MUTATION_FILES ?? '')
  .split(',')
  .filter((file) => file.length > 0);
const reportPath = process.env.MUTATION_REPORT;

if (mutate.length === 0) {
  throw new Error('MUTATION_FILES must select at least one source file');
}
if (reportPath === undefined || reportPath.length === 0) {
  throw new Error('MUTATION_REPORT must select a JSON report path');
}

export default {
  ...baseConfig,
  incremental: false,
  jsonReporter: { fileName: reportPath },
  mutate,
  reporters: ['json'],
};

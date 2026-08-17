import process from 'node:process';

import { resolveMutationModule } from './scripts/mutation-modules.mjs';
import baseConfig from './stryker.config.mjs';

const module = resolveMutationModule(process.env.MUTATION_MODULE ?? '');
const reportPath =
  process.env.MUTATION_REPORT ?? `reports/mutation/modules/${module.name}.json`;
process.env.MUTATION_TEST_FILES = module.tests.join(',');

export default {
  ...baseConfig,
  incremental: true,
  incrementalFile: `reports/mutation/cache/${module.name}.json`,
  jsonReporter: { fileName: reportPath },
  mutate: [module.source],
  reporters: ['clear-text', 'progress', 'json'],
  vitest: {
    ...(baseConfig.vitest ?? {}),
    configFile: 'vitest.mutation.config.ts',
    related: false,
  },
};

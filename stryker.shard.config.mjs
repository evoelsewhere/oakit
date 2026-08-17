import process from 'node:process';

import { mutationShardEnvironment } from './scripts/mutation-shard-environment.mjs';
import baseConfig from './stryker.config.mjs';

const { excludedMutations, mutate, reportPath } = mutationShardEnvironment(
  process.env,
);
const focusedTests = (process.env.MUTATION_TEST_FILES ?? '')
  .split(',')
  .filter((file) => file.length > 0);

export default {
  ...baseConfig,
  incremental: false,
  jsonReporter: { fileName: reportPath },
  mutate,
  mutator: {
    ...(baseConfig.mutator ?? {}),
    excludedMutations,
  },
  reporters: ['json'],
  ...(focusedTests.length === 0
    ? {}
    : {
        vitest: {
          ...(baseConfig.vitest ?? {}),
          configFile: 'vitest.mutation.config.ts',
          related: false,
        },
      }),
};

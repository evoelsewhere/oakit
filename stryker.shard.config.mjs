import process from 'node:process';

import { mutationShardEnvironment } from './scripts/mutation-shard-environment.mjs';
import baseConfig from './stryker.config.mjs';

const { excludedMutations, mutate, reportPath } = mutationShardEnvironment(
  process.env,
);

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
};

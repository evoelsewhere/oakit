import process from 'node:process';

import { defineConfig } from 'vitest/config';

const include = (process.env.MUTATION_TEST_FILES ?? '')
  .split(',')
  .filter((file) => file.length > 0);

if (include.length === 0) {
  throw new Error('MUTATION_TEST_FILES must select focused mutation tests');
}

export default defineConfig({
  test: {
    fileParallelism: false,
    include,
  },
});

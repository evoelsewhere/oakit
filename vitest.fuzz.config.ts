import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    fileParallelism: false,
    include: ['test/fuzz/**/*.test.ts'],
    outputFile: 'reports/reliability/pptx-render-fuzz.json',
    reporters: ['default', 'json'],
    testTimeout: 90_000,
  },
});

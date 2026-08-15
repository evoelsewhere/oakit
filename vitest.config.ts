import { configDefaults, defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    exclude: [
      ...configDefaults.exclude,
      '.stryker-tmp/**',
      '**/test/browser/**',
      '**/test/corpus/**',
    ],
  },
});

import { configDefaults, defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    bail: 1,
    exclude: [
      ...configDefaults.exclude,
      '.stryker-tmp/**',
      '**/test/browser/**',
      '**/test/corpus/**',
      '**/test/fuzz/**',
    ],
  },
});

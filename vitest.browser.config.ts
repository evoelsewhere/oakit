import { defineConfig } from 'vitest/config';

export default defineConfig({
  optimizeDeps: {
    include: ['txml/txml'],
  },
  test: {
    include: ['test/browser/**/*.test.ts'],
    browser: {
      enabled: true,
      headless: true,
      instances: [{ browser: 'chromium' }],
      provider: 'playwright',
    },
  },
});

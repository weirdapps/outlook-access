import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['test_scripts/**/*.spec.ts'],
    environment: 'node',
    testTimeout: 10000,
  },
});

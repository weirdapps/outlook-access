import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['test_scripts/**/*.spec.ts'],
    environment: 'node',
    testTimeout: 10000,
    coverage: {
      provider: 'v8',
      // SonarCloud reads coverage/lcov.info; vitest's default reporters omit
      // lcov, so the scan saw 0% coverage on new code and failed the gate.
      reporter: ['text', 'lcov'],
      reportsDirectory: 'coverage',
    },
  },
});

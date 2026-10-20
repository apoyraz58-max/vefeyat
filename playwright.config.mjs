import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './browser-tests',
  outputDir: './.local/test-results',
  fullyParallel: false,
  workers: 1,
  // Include system Chromium startup on shared cloud CPUs in the test budget.
  // Functional assertions keep Playwright's normal five-second deadline.
  timeout: 60_000,
  reporter: 'list',
  use: {
    baseURL: process.env.VEFEYAT_TEST_BASE_URL || 'http://127.0.0.1:8000',
    viewport: { width: 1440, height: 1000 },
    launchOptions: { executablePath: process.env.CHROMIUM_PATH || '/usr/bin/chromium' },
    screenshot: 'only-on-failure',
    trace: 'retain-on-failure',
  },
});

// Browser tests of the built page (npm run test:e2e): real Edge, Moodle answered by a fake.
import { defineConfig } from "vitest/config";

export default defineConfig({
  test: { include: ["test/e2e/**/*.e2e.ts"], testTimeout: 120_000, hookTimeout: 180_000, fileParallelism: false },
});

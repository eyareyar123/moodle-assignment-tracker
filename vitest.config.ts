// Unit tests (npm test). Kept apart from vite.config.ts, whose root is the web/ folder.
import { defineConfig } from "vitest/config";

export default defineConfig({
  test: { include: ["test/**/*.test.ts"] },
});

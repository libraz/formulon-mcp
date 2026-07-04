import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    // Tests exercise the built output (dist), so `yarn test` builds first.
    include: ["test/**/*.test.js"],
    environment: "node",
    globals: false,
    testTimeout: 30_000,
    coverage: {
      provider: "v8",
      reporter: ["text", "html", "lcov"],
      include: ["src/**/*.ts"],
      exclude: ["src/index.ts"],
    },
  },
});

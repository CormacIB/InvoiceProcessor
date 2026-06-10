import { defineConfig } from "vitest/config";
import react from "@vitejs/plugin-react";

export default defineConfig(({ mode }) => ({
  plugins: [react()],
  resolve: {
    alias:
      mode === "test"
        ? // pdf.js's standard build targets browsers; the legacy build runs in Node
          { "pdfjs-dist": "pdfjs-dist/legacy/build/pdf.mjs" }
        : {},
  },
  test: {
    environment: "node",
    testTimeout: 30000,
  },
}));

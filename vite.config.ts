/// <reference types="vitest/config" />
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

// Office add-ins must be served over HTTPS. `npm run certs` installs a trusted
// localhost certificate into ~/.office-addin-dev-certs. If it is missing we fall
// back to plain HTTP, which is still fine for the in-browser demo mode.
const certDir = join(homedir(), ".office-addin-dev-certs");
const certFile = join(certDir, "localhost.crt");
const keyFile = join(certDir, "localhost.key");
const https =
  existsSync(certFile) && existsSync(keyFile)
    ? { cert: readFileSync(certFile), key: readFileSync(keyFile) }
    : undefined;

export default defineConfig({
  plugins: [react()],
  root: "src/taskpane",
  publicDir: "../../assets",
  build: {
    outDir: "../../dist",
    emptyOutDir: true,
  },
  server: {
    port: 3000,
    strictPort: true,
    https,
  },
  test: {
    root: ".",
    include: ["tests/**/*.test.ts", "tests/**/*.test.tsx"],
    globals: true,
    environment: "node",
  },
});

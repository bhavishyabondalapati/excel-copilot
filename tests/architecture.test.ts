// Guards the core design rule: src/core is pure TypeScript. If someone imports
// Office.js, React or the network layer into it, this test fails.
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

const coreDir = join(__dirname, "..", "src", "core");

describe("src/core purity", () => {
  const files = readdirSync(coreDir).filter((f) => f.endsWith(".ts"));

  it.each(files)("%s has no Excel, React, DOM or SDK dependencies", (file) => {
    const src = readFileSync(join(coreDir, file), "utf8");
    const imports = [...src.matchAll(/from\s+"([^"]+)"/g)].map((m) => m[1]!);
    for (const spec of imports) {
      const allowed = spec.startsWith("./") || spec === "zod" || spec === "../providers/types";
      expect(allowed, `${file} imports "${spec}"`).toBe(true);
    }
    expect(src).not.toMatch(/\b(Excel|Office)\.\w/);
    expect(src).not.toMatch(/\b(document|window|localStorage)\./);
  });
});

import { readFileSync } from "node:fs";

const checks = [
  {
    file: "src/db.ts",
    expected: 'const DB_NAME = "research-supporter-v2";',
    label: "Predict IndexedDB name",
  },
  {
    file: "vite.config.js",
    expected: 'base: "/researchsupporter/",',
    label: "Predict GitHub Pages base path",
  },
];

const failures = checks.filter(({ file, expected }) => !readFileSync(file, "utf8").includes(expected));

if (failures.length > 0) {
  failures.forEach(({ file, expected, label }) => {
    console.error(`[invariant] ${label} was changed in ${file}. Expected: ${expected}`);
  });
  process.exit(1);
}

console.log("Predict repository invariants verified.");

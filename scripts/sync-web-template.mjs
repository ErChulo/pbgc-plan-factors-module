import { copyFile, mkdir } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const scriptDir = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(scriptDir, "..");

const source = resolve(repoRoot, "templates", "PF_template.v0.7.13.xlsx");
const targetDir = resolve(repoRoot, "packages", "web-pf", "public");
const target = resolve(targetDir, "PF_template.v0.7.13.xlsx");

try {
  await mkdir(targetDir, { recursive: true });
  await copyFile(source, target);
  console.log(`Copied template to ${target}`);
} catch (error) {
  const code = error && typeof error === "object" && "code" in error ? error.code : undefined;
  if (code === "ENOENT") {
    console.warn("Template not found at templates/PF_template.v0.7.13.xlsx. Add it to enable web generation.");
  } else {
    throw error;
  }
}

#!/usr/bin/env node
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { basename, join, resolve } from "node:path";
import { buildPfWorkbook, validateCaseJson, validatePlanFactorsJson } from "@pbgc/core-pf";

type ArgMap = Record<string, string | undefined>;

function parseArgs(argv: string[]): ArgMap {
  const map: ArgMap = {};
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg.startsWith("--")) {
      map[arg.slice(2)] = argv[i + 1];
      i += 1;
    }
  }
  return map;
}

async function run(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));
  const casePath = args.case;
  const planPath = args.plan;
  const outDir = args.out ?? ".";
  const templatePath = args.template ?? "templates/PF_template.v0.7.13.xlsx";

  if (!casePath || !planPath) {
    throw new Error("Usage: pf-build --case case.json --plan planFactors.json --out outdir [--template templates/PF_template.v0.7.13.xlsx]");
  }

  const [caseRaw, planRaw, templateBytes] = await Promise.all([
    readFile(resolve(casePath), "utf8"),
    readFile(resolve(planPath), "utf8"),
    readFile(resolve(templatePath))
  ]);

  const caseInput = JSON.parse(caseRaw);
  const planFactorsInput = JSON.parse(planRaw);

  const caseValidation = validateCaseJson(caseInput);
  const planValidation = validatePlanFactorsJson(planFactorsInput);
  const errors = [
    ...caseValidation.errors.map((e: { path: string; message: string }) => `case.json ${e.path}: ${e.message}`),
    ...planValidation.errors.map((e: { path: string; message: string }) => `planFactors.json ${e.path}: ${e.message}`)
  ];

  if (errors.length > 0) {
    throw new Error(`Validation failed:\n${errors.join("\n")}`);
  }

  const built = await buildPfWorkbook({
    caseInput,
    planFactorsInput,
    templateBytes: new Uint8Array(templateBytes)
  });

  await mkdir(resolve(outDir), { recursive: true });
  const outPath = join(resolve(outDir), basename(built.fileName));
  await writeFile(outPath, built.bytes);
  process.stdout.write(`Generated: ${outPath}\n`);
}

run().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
  process.exit(1);
});
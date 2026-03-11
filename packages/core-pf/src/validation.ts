import Ajv2020 from "ajv/dist/2020.js";
import addFormats from "ajv-formats";
import caseSchema from "../../schemas/case.schema.json" with { type: "json" };
import planFactorsSchema from "../../schemas/planFactors.schema.json" with { type: "json" };
import type { CaseInput, PlanFactorsInput, ValidationResult } from "./types.js";

const ajv = new Ajv2020({ allErrors: true, strict: false });
addFormats(ajv);

const validateCaseInput = ajv.compile(caseSchema);
const validatePlanFactors = ajv.compile(planFactorsSchema);

function toErrors(errors: typeof validateCaseInput.errors): { path: string; message: string }[] {
  return (errors ?? []).map((err) => ({
    path: err.instancePath || "/",
    message: err.message ?? "Invalid value"
  }));
}

export function validateCaseJson(value: unknown): ValidationResult<CaseInput> {
  const valid = validateCaseInput(value);
  if (!valid) {
    return { valid: false, errors: toErrors(validateCaseInput.errors) };
  }
  return { valid: true, errors: [], value: value as unknown as CaseInput };
}

export function validatePlanFactorsJson(value: unknown): ValidationResult<PlanFactorsInput> {
  const valid = validatePlanFactors(value);
  if (!valid) {
    return { valid: false, errors: toErrors(validatePlanFactors.errors) };
  }
  return { valid: true, errors: [], value: value as unknown as PlanFactorsInput };
}

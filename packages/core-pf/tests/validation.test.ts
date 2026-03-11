import { describe, expect, it } from "vitest";
import { validateCaseJson, validatePlanFactorsJson } from "../src/validation.js";
import caseSample from "../../../fixtures/case.sample.json" with { type: "json" };
import planSample from "../../../fixtures/planFactors.sample.json" with { type: "json" };

describe("schema validation", () => {
  it("accepts sample case json", () => {
    const result = validateCaseJson(caseSample);
    expect(result.valid).toBe(true);
  });

  it("accepts sample plan factors json", () => {
    const result = validatePlanFactorsJson(planSample);
    expect(result.valid).toBe(true);
  });

  it("rejects invalid case number", () => {
    const result = validateCaseJson({ ...caseSample, caseNumber: "ABC" });
    expect(result.valid).toBe(false);
  });
});
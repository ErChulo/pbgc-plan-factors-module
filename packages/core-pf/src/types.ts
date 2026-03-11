import { PF_VERSION } from "./constants.js";

export type Sex = "M" | "F";
export type DeferralMortalityFlag = "Y" | "N";

export interface CaseInput {
  caseNumber: string;
  planName: string;
  dateOfPlanTermination: string;
  normalRetirementAge: number;
}

export interface ActuarialFactor {
  interest: number;
  sex: Sex;
  monthsCertain: number;
  mortalityMale: string;
  mortalityFemale: string;
  calcMethod: "MP";
  deferralMortality: DeferralMortalityFlag;
}

export interface BfcfFactor {
  fromFormAbbr: "SLA" | "3CC";
  fromMonthsCertain: number;
  toFormAbbr: "JS50";
  toSurvivorPercent: 0.5;
  interest: number;
  mortalityMale: string;
  mortalityFemale: string;
  annuityType: "MP";
  method: "N";
}

export interface PlanFactorsInput {
  earlyRetirement: {
    "ERF-1976": { type: "linear_1_over_15_per_year" };
    "ERF-1979": { type: "tiered_monthly_1_180_then_1_360" };
    "ERF-Disability": ActuarialFactor & { type?: string };
  };
  lateRetirement: {
    "LRF-1978": ActuarialFactor;
    "LRF-1982": ActuarialFactor;
    "LRF-2011": ActuarialFactor;
  };
  benefitFormConversion: {
    "BFCF-1976": BfcfFactor;
    "BFCF-1979": BfcfFactor;
    "BFCF-1997": BfcfFactor;
    "BFCF-2011": BfcfFactor;
    "BFCF-2019": BfcfFactor;
  };
  normalSingleForm: {
    "1976": "SLA" | "3CC";
    "1979": "SLA" | "3CC";
    "1997": "SLA" | "3CC";
    "2011": "SLA" | "3CC";
    "2019": "SLA" | "3CC";
  };
}

export interface ValidationIssue {
  path: string;
  message: string;
}

export interface ValidationResult<T> {
  valid: boolean;
  errors: ValidationIssue[];
  value?: T;
}

export interface BuildPfWorkbookInput {
  caseInput: CaseInput;
  planFactorsInput: PlanFactorsInput;
  templateBytes: Uint8Array;
}

export interface BuildPfWorkbookResult {
  bytes: Uint8Array;
  fileName: string;
  version: typeof PF_VERSION;
}
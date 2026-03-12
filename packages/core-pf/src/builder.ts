import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";
import { PF_SHEET_ORDER, PF_VERSION } from "./constants.js";
import type { PfSheetName } from "./constants.js";
import type {
  BfcfFactor,
  BuildPfWorkbookInput,
  BuildPfWorkbookResult,
  CaseInput,
  PlanFactorsInput
} from "./types.js";
import {
  formatDateMmDdYyyy,
  getLeftAdjacentCell,
  iterateRange,
  parseWorksheetXml,
  serializeXml,
  setFormulaCell,
  setInlineStringCell
} from "./ooxml.js";
import { validateCaseJson, validatePlanFactorsJson } from "./validation.js";

function quote(value: string): string {
  return `"${value}"`;
}

function toRelPath(target: string): string {
  return target.replace(/^\/?xl\//, "");
}

function fmtPercent(value: number, digits = 0): string {
  return `${(value * 100).toFixed(digits)}%`;
}

async function getSheetFileMap(zip: JSZip): Promise<Record<string, string>> {
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");

  if (!workbookXml || !relsXml) {
    throw new Error("Template workbook is missing workbook relationship files.");
  }

  const parser = new DOMParser();
  const workbookDoc = parser.parseFromString(workbookXml, "text/xml");
  const relsDoc = parser.parseFromString(relsXml, "text/xml");

  const relMap: Record<string, string> = {};
  Array.from(relsDoc.getElementsByTagName("Relationship")).forEach((rel) => {
    const id = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    if (id && target) relMap[id] = toRelPath(target);
  });

  const map: Record<string, string> = {};
  Array.from(workbookDoc.getElementsByTagName("sheet")).forEach((sheet) => {
    const name = sheet.getAttribute("name");
    const rId = sheet.getAttribute("r:id");
    if (!name || !rId) return;

    const relPath = relMap[rId];
    if (relPath) {
      map[name] = `xl/${relPath}`;
    }
  });

  return map;
}

function erfFormula(
  sheetName: PfSheetName,
  monthsCell: string,
  nraCell = "65",
  planFactors?: PlanFactorsInput
): string {
  if (sheetName === "ERF-1976") {
    return `=ROUND(1-(${monthsCell}/180),4)`;
  }

  if (sheetName === "ERF-1979") {
    return `=ROUND(1-(1/180)*MIN(${monthsCell},60)-(1/360)*MAX(MIN(${monthsCell}-60,60),0),4)`;
  }

  const disability = planFactors?.earlyRetirement["ERF-Disability"];
  if (!disability) {
    throw new Error("Missing ERF-Disability assumptions");
  }

  return `=ROUND(ERFAEQ(${disability.interest},${quote("M")},${nraCell},${monthsCell},${disability.monthsCertain},${quote(disability.mortalityMale)},${quote(disability.mortalityFemale)},${quote("MP")},${quote(disability.deferralMortality)}),4)`;
}

function lrfFormula(
  monthsCell: string,
  assumptions: PlanFactorsInput["lateRetirement"]["LRF-1978"],
  nraCell = "65"
): string {
  return `=ROUND(LRFAEQ(${assumptions.interest},${quote("M")},${nraCell},${monthsCell},${assumptions.monthsCertain},${quote(assumptions.mortalityMale)},${quote(assumptions.mortalityFemale)},${quote("MP")},${quote(assumptions.deferralMortality)}),4)`;
}

function bfcfFormula(cellRef: string, assumptions: BfcfFactor): string {
  const participantAgeCell = `${cellRef.replace(/[0-9]+$/, "")}10`;
  const beneficiaryAgeCell = `B${cellRef.match(/[0-9]+$/)?.[0]}`;

  return `=ROUND(BFCFAEQ(${quote(assumptions.fromFormAbbr)},0,${assumptions.fromMonthsCertain},0,${quote("JSC")},0.5,0,0,${participantAgeCell},${beneficiaryAgeCell},${quote("M")},${quote("M")},${assumptions.interest},${quote(assumptions.mortalityMale)},${quote(assumptions.mortalityFemale)},${quote("MP")},${quote("N")}),4)`;
}

function setCommonHeader(caseInput: CaseInput, doc: Document, worksheet: Element, baseCell: "A" | "B"): void {
  setInlineStringCell(doc, worksheet, `${baseCell}1`, caseInput.planName);
  setInlineStringCell(doc, worksheet, `${baseCell}2`, `Case Number: ${caseInput.caseNumber}`);
  setInlineStringCell(doc, worksheet, `${baseCell}3`, `DOPT: ${formatDateMmDdYyyy(caseInput.dateOfPlanTermination)}`);
}

function normalSingleFormForSheet(sheetName: PfSheetName, normalSingleForm: PlanFactorsInput["normalSingleForm"]): string {
  if (sheetName === "ERF-1976") return normalSingleForm["1976"];
  if (sheetName === "ERF-1979") return normalSingleForm["1979"];
  if (sheetName === "ERF-Disability") return normalSingleForm["2011"];
  if (sheetName === "LRF-1978") return normalSingleForm["1976"];
  if (sheetName === "LRF-1982") return normalSingleForm["1997"];
  if (sheetName === "LRF-2011") return normalSingleForm["2011"];
  if (sheetName === "BFCF-1976 SLA->JS50") return normalSingleForm["1976"];
  if (sheetName === "BFCF-1979 SLA->JS50") return normalSingleForm["1979"];
  if (sheetName === "BFCF-1997 3CC->JS50") return normalSingleForm["1997"];
  if (sheetName === "BFCF-2011 3CC->JS50") return normalSingleForm["2011"];
  return normalSingleForm["2019"];
}

function patchErfSheet(
  sheetName: "ERF-1976" | "ERF-1979" | "ERF-Disability",
  caseInput: CaseInput,
  planFactors: PlanFactorsInput,
  doc: Document,
  worksheet: Element
): void {
  setCommonHeader(caseInput, doc, worksheet, "A");

  if (sheetName === "ERF-1976") {
    setInlineStringCell(doc, worksheet, "A5", `Early Retirement Factors from NRA of ${caseInput.normalRetirementAge} (1976 Plan)`);
    setInlineStringCell(doc, worksheet, "A6", "Basis: reduction = 1/15 per year early (linear in months)");
  } else if (sheetName === "ERF-1979") {
    setInlineStringCell(doc, worksheet, "A5", `Early Retirement Factors from NRA of ${caseInput.normalRetirementAge} (1979/1997/2011/2019 Plans)`);
    setInlineStringCell(doc, worksheet, "A6", "Basis: reduction = 1/180 per month for first 60 months, then 1/360 per month for next 60 months");
  } else {
    const d = planFactors.earlyRetirement["ERF-Disability"];
    setInlineStringCell(doc, worksheet, "A5", `Early Retirement Factors from NRA of ${caseInput.normalRetirementAge} (Disability (2011/2019))`);
    setInlineStringCell(
      doc,
      worksheet,
      "A6",
      `Basis: ERFAEQ with interest=${fmtPercent(d.interest, 0)}; mortality=${d.mortalityMale}/${d.mortalityFemale}; calc method=MP; deferral mortality=${d.deferralMortality}; months certain=${d.monthsCertain}`
    );
  }

  setInlineStringCell(doc, worksheet, "A7", `Normal single form: ${normalSingleFormForSheet(sheetName as PfSheetName, planFactors.normalSingleForm)}`);

  for (let row = 11; row <= 34; row += 1) {
    for (const col of ["B", "D", "F", "H", "J"]) {
      const target = `${col}${row}`;
      const monthsCell = getLeftAdjacentCell(target);
      setFormulaCell(doc, worksheet, target, erfFormula(sheetName, monthsCell, String(caseInput.normalRetirementAge), planFactors));
    }
  }
}

function patchLrfSheet(
  sheetName: "LRF-1978" | "LRF-1982" | "LRF-2011",
  caseInput: CaseInput,
  planFactors: PlanFactorsInput,
  doc: Document,
  worksheet: Element
): void {
  setCommonHeader(caseInput, doc, worksheet, "A");
  const assumptions = planFactors.lateRetirement[sheetName];

  if (sheetName === "LRF-1978") {
    setInlineStringCell(doc, worksheet, "A5", `Late Retirement Factors from NRA of ${caseInput.normalRetirementAge} (1978 Plan)`);
  } else if (sheetName === "LRF-1982") {
    setInlineStringCell(doc, worksheet, "A5", `Late Retirement Factors from NRA of ${caseInput.normalRetirementAge} (1982 Plan)`);
  } else {
    setInlineStringCell(doc, worksheet, "A5", `Late Retirement Factors from NRA of ${caseInput.normalRetirementAge} (2011/2019 Plans)`);
  }

  setInlineStringCell(
    doc,
    worksheet,
    "A6",
    `Basis: LRFAEQ with interest=${fmtPercent(assumptions.interest, 0)}; mortality=${assumptions.mortalityMale}/${assumptions.mortalityFemale}; calc method=MP; deferral mortality=${assumptions.deferralMortality}; months certain=${assumptions.monthsCertain}`
  );

  setInlineStringCell(doc, worksheet, "A7", `Normal single form: ${normalSingleFormForSheet(sheetName as PfSheetName, planFactors.normalSingleForm)}`);

  for (let row = 11; row <= 34; row += 1) {
    for (const col of ["B", "D", "F", "H", "J"]) {
      const target = `${col}${row}`;
      const monthsCell = getLeftAdjacentCell(target);
      setFormulaCell(doc, worksheet, target, lrfFormula(monthsCell, assumptions, String(caseInput.normalRetirementAge)));
    }
  }
}

const BFCF_SHEET_TO_KEY: Record<string, keyof PlanFactorsInput["benefitFormConversion"]> = {
  "BFCF-1976 SLA->JS50": "BFCF-1976",
  "BFCF-1979 SLA->JS50": "BFCF-1979",
  "BFCF-1997 3CC->JS50": "BFCF-1997",
  "BFCF-2011 3CC->JS50": "BFCF-2011",
  "BFCF-2019 3CC->JS50": "BFCF-2019"
};

const BFCF_SHEET_TO_PLAN_LABEL: Record<string, string> = {
  "BFCF-1976 SLA->JS50": "1976 Plan",
  "BFCF-1979 SLA->JS50": "1979 Plan",
  "BFCF-1997 3CC->JS50": "1997 Plan",
  "BFCF-2011 3CC->JS50": "2011 Plan",
  "BFCF-2019 3CC->JS50": "2019 Plan"
};

function patchBfcfSheet(
  sheetName: keyof typeof BFCF_SHEET_TO_KEY,
  caseInput: CaseInput,
  planFactors: PlanFactorsInput,
  doc: Document,
  worksheet: Element
): void {
  setCommonHeader(caseInput, doc, worksheet, "B");
  const key = BFCF_SHEET_TO_KEY[sheetName];
  const assumptions = planFactors.benefitFormConversion[key];

  setInlineStringCell(
    doc,
    worksheet,
    "B5",
    `Form Conversion Factors: ${assumptions.fromFormAbbr}->JS50 (${BFCF_SHEET_TO_PLAN_LABEL[sheetName as keyof typeof BFCF_SHEET_TO_PLAN_LABEL]})`
  );
  setInlineStringCell(
    doc,
    worksheet,
    "B6",
    `Basis: interest=${fmtPercent(assumptions.interest, 2)}; mortality=${assumptions.mortalityMale}/${assumptions.mortalityFemale}; method=MP`
  );
  setInlineStringCell(doc, worksheet, "B7", `Normal single form: ${normalSingleFormForSheet(sheetName as PfSheetName, planFactors.normalSingleForm)}`);

  const cells = iterateRange("C", "AB", 11, 111);
  for (const cellRef of cells) {
    setFormulaCell(doc, worksheet, cellRef, bfcfFormula(cellRef, assumptions));
  }
}

export async function buildPfWorkbook(input: BuildPfWorkbookInput): Promise<BuildPfWorkbookResult> {
  const caseValidation = validateCaseJson(input.caseInput);
  if (!caseValidation.valid) {
    throw new Error(`Invalid case input: ${caseValidation.errors.map((e) => `${e.path} ${e.message}`).join("; ")}`);
  }

  const planValidation = validatePlanFactorsJson(input.planFactorsInput);
  if (!planValidation.valid) {
    throw new Error(`Invalid plan factors input: ${planValidation.errors.map((e) => `${e.path} ${e.message}`).join("; ")}`);
  }

  const zip = await JSZip.loadAsync(input.templateBytes);
  const sheetMap = await getSheetFileMap(zip);

  for (const sheetName of PF_SHEET_ORDER) {
    const sheetPath = sheetMap[sheetName];
    if (!sheetPath) {
      throw new Error(`Template missing expected sheet: ${sheetName}`);
    }

    const xml = await zip.file(sheetPath)?.async("string");
    if (!xml) {
      throw new Error(`Worksheet XML not found for ${sheetName}`);
    }

    const { doc, worksheet } = parseWorksheetXml(xml);

    if (sheetName === "ERF-1976" || sheetName === "ERF-1979" || sheetName === "ERF-Disability") {
      patchErfSheet(sheetName, input.caseInput, input.planFactorsInput, doc, worksheet);
    } else if (sheetName === "LRF-1978" || sheetName === "LRF-1982" || sheetName === "LRF-2011") {
      patchLrfSheet(sheetName, input.caseInput, input.planFactorsInput, doc, worksheet);
    } else {
      patchBfcfSheet(sheetName as keyof typeof BFCF_SHEET_TO_KEY, input.caseInput, input.planFactorsInput, doc, worksheet);
    }

    zip.file(sheetPath, serializeXml(doc));
  }

  const bytes = await zip.generateAsync({ type: "uint8array" });

  return {
    bytes,
    fileName: `${input.caseInput.caseNumber}PF.v${PF_VERSION}.xlsx`,
    version: PF_VERSION
  };
}

export function buildOutputFileName(caseNumber: string): string {
  return `${caseNumber}PF.v${PF_VERSION}.xlsx`;
}



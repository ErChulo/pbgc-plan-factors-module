import JSZip from "jszip";
import { DOMParser } from "@xmldom/xmldom";
import { HEADER_CELL_MAP, PF_SHEET_ORDER, PF_VERSION } from "./constants.js";
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
  nraCell = "$H$3",
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
  nraCell = "$H$3"
): string {
  return `=ROUND(LRFAEQ(${assumptions.interest},${quote("M")},${nraCell},${monthsCell},${assumptions.monthsCertain},${quote(assumptions.mortalityMale)},${quote(assumptions.mortalityFemale)},${quote("MP")},${quote(assumptions.deferralMortality)}),4)`;
}

function bfcfFormula(cellRef: string, assumptions: BfcfFactor): string {
  const participantAgeCell = `${cellRef.replace(/[0-9]+$/, "")}10`;
  const beneficiaryAgeCell = `B${cellRef.match(/[0-9]+$/)?.[0]}`;

  return `=ROUND(BFCFAEQ(${quote("SLA")},0,${assumptions.fromMonthsCertain},0,${quote("JSC")},0.5,0,0,${participantAgeCell},${beneficiaryAgeCell},${quote("M")},${quote("M")},${assumptions.interest},${quote(assumptions.mortalityMale)},${quote(assumptions.mortalityFemale)},${quote("MP")},${quote("N")}),4)`;
}

function patchHeaderCells(caseInput: CaseInput, doc: Document, worksheet: Element): void {
  setInlineStringCell(doc, worksheet, HEADER_CELL_MAP.caseNumber, caseInput.caseNumber);
  setInlineStringCell(doc, worksheet, HEADER_CELL_MAP.planName, caseInput.planName);
  setInlineStringCell(doc, worksheet, HEADER_CELL_MAP.dopt, formatDateMmDdYyyy(caseInput.dateOfPlanTermination));
  setInlineStringCell(doc, worksheet, HEADER_CELL_MAP.nra, String(caseInput.normalRetirementAge));
}

function patchErfSheet(
  sheetName: PfSheetName,
  caseInput: CaseInput,
  planFactors: PlanFactorsInput,
  doc: Document,
  worksheet: Element
): void {
  patchHeaderCells(caseInput, doc, worksheet);

  for (let row = 11; row <= 34; row += 1) {
    for (const col of ["B", "D", "F", "H", "J"]) {
      const target = `${col}${row}`;
      const monthsCell = getLeftAdjacentCell(target);
      setFormulaCell(doc, worksheet, target, erfFormula(sheetName, monthsCell, "$H$3", planFactors));
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
  patchHeaderCells(caseInput, doc, worksheet);
  const assumptions = planFactors.lateRetirement[sheetName];

  for (let row = 11; row <= 34; row += 1) {
    for (const col of ["B", "D", "F", "H", "J"]) {
      const target = `${col}${row}`;
      const monthsCell = getLeftAdjacentCell(target);
      setFormulaCell(doc, worksheet, target, lrfFormula(monthsCell, assumptions));
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

function patchBfcfSheet(
  sheetName: string,
  caseInput: CaseInput,
  planFactors: PlanFactorsInput,
  doc: Document,
  worksheet: Element
): void {
  patchHeaderCells(caseInput, doc, worksheet);
  const assumptions = planFactors.benefitFormConversion[BFCF_SHEET_TO_KEY[sheetName]];

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

    if (sheetName.startsWith("ERF")) {
      patchErfSheet(sheetName, input.caseInput, input.planFactorsInput, doc, worksheet);
    } else if (sheetName.startsWith("LRF")) {
      patchLrfSheet(sheetName as "LRF-1978" | "LRF-1982" | "LRF-2011", input.caseInput, input.planFactorsInput, doc, worksheet);
    } else {
      patchBfcfSheet(sheetName, input.caseInput, input.planFactorsInput, doc, worksheet);
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

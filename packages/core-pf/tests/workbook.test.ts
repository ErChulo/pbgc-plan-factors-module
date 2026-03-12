import { createHash } from "node:crypto";
import { access, readFile } from "node:fs/promises";
import { constants } from "node:fs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import caseSample from "../../../fixtures/case.sample.json" with { type: "json" };
import planSample from "../../../fixtures/planFactors.sample.json" with { type: "json" };
import { PF_SHEET_ORDER, buildPfWorkbook } from "../src/index.js";

const TEMPLATE_PATH = "templates/PF_template.v0.7.13.xlsx";
const EXPECTED_PATH = "fixtures/expected.24884900PF.v0.7.13.xlsx";

async function exists(path: string): Promise<boolean> {
  try {
    await access(path, constants.F_OK);
    return true;
  } catch {
    return false;
  }
}

function parseSheetNames(workbookXml: string): string[] {
  return Array.from(workbookXml.matchAll(/<sheet[^>]*name="([^"]+)"/g)).map((m) => m[1]);
}

function parseRelationships(relsXml: string): Record<string, string> {
  const map: Record<string, string> = {};
  for (const m of relsXml.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    map[m[1]] = m[2].replace(/^\/?xl\//, "");
  }
  return map;
}

function parseSheetMap(workbookXml: string, relsMap: Record<string, string>): Record<string, string> {
  const map: Record<string, string> = {};
  for (const m of workbookXml.matchAll(/<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)) {
    const name = m[1];
    const rid = m[2];
    const target = relsMap[rid];
    if (target) map[name] = `xl/${target}`;
  }
  return map;
}

describe("workbook generation", () => {
  it("matches golden workbook hash when fixtures are present", async () => {
    const hasTemplate = await exists(TEMPLATE_PATH);
    const hasGolden = await exists(EXPECTED_PATH);

    if (!hasTemplate || !hasGolden) return;

    const [template, expected] = await Promise.all([readFile(TEMPLATE_PATH), readFile(EXPECTED_PATH)]);

    const built = await buildPfWorkbook({
      caseInput: caseSample,
      planFactorsInput: planSample,
      templateBytes: new Uint8Array(template)
    });

    const generatedHash = createHash("sha256").update(Buffer.from(built.bytes)).digest("hex");
    const expectedHash = createHash("sha256").update(expected).digest("hex");
    expect(generatedHash).toBe(expectedHash);
  });

  it("keeps required sheet order and workbook defined names", async () => {
    const hasTemplate = await exists(TEMPLATE_PATH);
    if (!hasTemplate) return;

    const template = await readFile(TEMPLATE_PATH);
    const built = await buildPfWorkbook({
      caseInput: caseSample,
      planFactorsInput: planSample,
      templateBytes: new Uint8Array(template)
    });

    const zip = await JSZip.loadAsync(built.bytes);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
    expect(workbookXml).toBeTruthy();

    const sheetOrder = parseSheetNames(workbookXml ?? "");
    expect(sheetOrder).toEqual([...PF_SHEET_ORDER]);

    expect((workbookXml ?? "").includes("_xlnm.Print_Area")).toBe(true);
    expect((workbookXml ?? "").includes("_xlnm.Print_Titles")).toBe(true);
    expect((workbookXml ?? "").includes("'ERF-1976'!$A$1:$J$34")).toBe(true);
    expect((workbookXml ?? "").includes("'BFCF-1976 SLA-&gt;JS50'!$A$1:$AB$111")).toBe(true);
  });

  it("preserves freeze panes and page setup", async () => {
    const hasTemplate = await exists(TEMPLATE_PATH);
    if (!hasTemplate) return;

    const template = await readFile(TEMPLATE_PATH);
    const built = await buildPfWorkbook({
      caseInput: caseSample,
      planFactorsInput: planSample,
      templateBytes: new Uint8Array(template)
    });

    const zip = await JSZip.loadAsync(built.bytes);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    const relsMap = parseRelationships(relsXml ?? "");
    const sheetMap = parseSheetMap(workbookXml ?? "", relsMap);

    for (const sheetName of PF_SHEET_ORDER) {
      const xml = (await zip.file(sheetMap[sheetName])?.async("string")) ?? "";
      if (sheetName.startsWith("ERF") || sheetName.startsWith("LRF")) {
        expect(xml.includes('topLeftCell="A11"')).toBe(true);
      } else {
        expect(xml.includes('topLeftCell="C11"')).toBe(true);
      }
      expect(xml.includes('orientation="landscape"')).toBe(true);
      expect(xml.includes('fitToWidth="1"')).toBe(true);
      expect(xml.includes('fitToHeight="2"')).toBe(true);
      expect(xml.includes('left="0.4"')).toBe(true);
    }
  });

  it("updates exact header cells and formulas for changed inputs", async () => {
    const hasTemplate = await exists(TEMPLATE_PATH);
    if (!hasTemplate) return;

    const template = await readFile(TEMPLATE_PATH);
    const caseInput = {
      ...caseSample,
      caseNumber: "99999999",
      planName: "Sample Custom Plan",
      dateOfPlanTermination: "2025-12-31",
      normalRetirementAge: 67
    };
    const planInput = {
      ...planSample,
      earlyRetirement: {
        ...planSample.earlyRetirement,
        "ERF-Disability": {
          ...planSample.earlyRetirement["ERF-Disability"],
          interest: 0.07,
          monthsCertain: 24,
          mortalityMale: "UP84",
          mortalityFemale: "UP84",
          deferralMortality: "N"
        }
      },
      benefitFormConversion: {
        ...planSample.benefitFormConversion,
        "BFCF-1997": {
          ...planSample.benefitFormConversion["BFCF-1997"],
          fromFormAbbr: "3CC",
          fromMonthsCertain: 36,
          interest: 0.09
        }
      }
    };

    const built = await buildPfWorkbook({
      caseInput,
      planFactorsInput: planInput,
      templateBytes: new Uint8Array(template)
    });

    const zip = await JSZip.loadAsync(built.bytes);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    const relsMap = parseRelationships(relsXml ?? "");
    const sheetMap = parseSheetMap(workbookXml ?? "", relsMap);

    const erf1976 = (await zip.file(sheetMap["ERF-1976"])?.async("string")) ?? "";
    expect(erf1976.includes("Sample Custom Plan")).toBe(true);
    expect(erf1976.includes("Case Number: 99999999")).toBe(true);
    expect(erf1976.includes("DOPT: 12/31/2025")).toBe(true);
    expect(erf1976.includes("NRA of 67")).toBe(true);

    const disability = (await zip.file(sheetMap["ERF-Disability"])?.async("string")) ?? "";
    expect(disability.includes("interest=7%")).toBe(true);
    expect(disability.includes("months certain=24")).toBe(true);
    expect(disability.includes("deferral mortality=N")).toBe(true);
    expect(disability.includes("ERFAEQ(0.07")).toBe(true);

    const bfcf1997 = (await zip.file(sheetMap["BFCF-1997 3CC->JS50"])?.async("string")) ?? "";
    expect(bfcf1997.includes("Form Conversion Factors: 3CC->JS50 (1997 Plan)")).toBe(true);
    expect(bfcf1997.includes("interest=9.00%")).toBe(true);
    expect(bfcf1997.includes('BFCFAEQ("3CC",0,36,0')).toBe(true);
  });
});

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

const ERF_LRF = new Set(["ERF-1976", "ERF-1979", "ERF-Disability", "LRF-1978", "LRF-1982", "LRF-2011"]);
const BFCF = new Set([
  "BFCF-1976 SLA->JS50",
  "BFCF-1979 SLA->JS50",
  "BFCF-1997 3CC->JS50",
  "BFCF-2011 3CC->JS50",
  "BFCF-2019 3CC->JS50"
]);

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

function find(xml: string, re: RegExp): string | null {
  const match = re.exec(xml);
  return match?.[1] ?? null;
}

describe("workbook generation", () => {
  it("matches golden workbook hash when fixtures are present", async () => {
    const hasTemplate = await exists(TEMPLATE_PATH);
    const hasGolden = await exists(EXPECTED_PATH);

    if (!hasTemplate || !hasGolden) {
      return;
    }

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

  it("keeps required sheet order", async () => {
    const hasTemplate = await exists(TEMPLATE_PATH);
    if (!hasTemplate) {
      return;
    }

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
  });

  it("preserves print settings and freeze panes", async () => {
    const hasTemplate = await exists(TEMPLATE_PATH);
    if (!hasTemplate) {
      return;
    }

    const template = await readFile(TEMPLATE_PATH);
    const built = await buildPfWorkbook({
      caseInput: caseSample,
      planFactorsInput: planSample,
      templateBytes: new Uint8Array(template)
    });

    const zip = await JSZip.loadAsync(built.bytes);
    const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    expect(workbookXml).toBeTruthy();
    expect(relsXml).toBeTruthy();

    const relsMap = parseRelationships(relsXml ?? "");
    const sheetMap = parseSheetMap(workbookXml ?? "", relsMap);

    for (const sheetName of PF_SHEET_ORDER) {
      const xml = await zip.file(sheetMap[sheetName])?.async("string");
      expect(xml).toBeTruthy();

      const pane = find(xml ?? "", /<pane[^>]*topLeftCell="([^"]+)"/);
      const printArea = find(xml ?? "", /<definedName[^>]*_xlnm\.Print_Area[^>]*>([^<]+)</);
      const printTitles = find(xml ?? "", /<definedName[^>]*_xlnm\.Print_Titles[^>]*>([^<]+)</);

      if (ERF_LRF.has(sheetName)) {
        expect(pane).toBe("A11");
      }
      if (BFCF.has(sheetName)) {
        expect(pane).toBe("C11");
      }

      if (ERF_LRF.has(sheetName)) {
        expect((xml ?? "").includes("A1:J34")).toBe(true);
      }
      if (BFCF.has(sheetName)) {
        expect((xml ?? "").includes("A1:AB111")).toBe(true);
      }

      expect((xml ?? "").includes("<pageSetup")).toBe(true);
      expect((xml ?? "").includes("orientation=\"landscape\"")).toBe(true);
      expect((xml ?? "").includes("fitToWidth=\"1\"")).toBe(true);
      expect((xml ?? "").includes("fitToHeight=\"2\"")).toBe(true);
      expect((xml ?? "").includes("<pageMargins")).toBe(true);
      expect((xml ?? "").includes("left=\"0.4\"")).toBe(true);

      expect(printArea).not.toBeNull();
      expect(printTitles).not.toBeNull();
    }
  });

  it("writes representative formulas", async () => {
    const hasTemplate = await exists(TEMPLATE_PATH);
    if (!hasTemplate) {
      return;
    }

    const template = await readFile(TEMPLATE_PATH);
    const built = await buildPfWorkbook({
      caseInput: caseSample,
      planFactorsInput: planSample,
      templateBytes: new Uint8Array(template)
    });

    const zip = await JSZip.loadAsync(built.bytes);
    const allSheets = Object.keys(zip.files).filter((name) => name.startsWith("xl/worksheets/sheet") && name.endsWith(".xml"));
    const joinedXml = (await Promise.all(allSheets.map(async (p) => zip.file(p)?.async("string") ?? ""))).join("\n");

    expect(joinedXml.includes("ROUND(")).toBe(true);
    expect(joinedXml.includes("LRFAEQ(")).toBe(true);
    expect(joinedXml.includes("BFCFAEQ(")).toBe(true);
    expect(joinedXml.includes("ERFAEQ(")).toBe(true);
  });
});

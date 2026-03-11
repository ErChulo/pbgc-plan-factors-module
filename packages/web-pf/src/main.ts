import "./styles.css";
import { PF_SHEET_ORDER, buildPfWorkbook, validateCaseJson, validatePlanFactorsJson, type CaseInput, type PlanFactorsInput } from "@pbgc/core-pf";

const app = document.querySelector<HTMLDivElement>("#app");
if (!app) throw new Error("#app not found");

app.innerHTML = `
  <main>
    <h1>Plan Factors Workbook Builder (v0.7.13)</h1>
    <small>Uploads stay in your browser. No server calls.</small>

    <div class="grid">
      <label>Case JSON
        <input id="caseFile" type="file" accept="application/json,.json" />
      </label>
      <label>Plan Factors JSON
        <input id="planFile" type="file" accept="application/json,.json" />
      </label>
    </div>

    <button id="generateBtn">Generate PF workbook</button>
    <div id="errors" class="errors"></div>
    <div id="summary" class="card" hidden></div>
  </main>
`;

const caseFileInput = document.querySelector<HTMLInputElement>("#caseFile")!;
const planFileInput = document.querySelector<HTMLInputElement>("#planFile")!;
const generateBtn = document.querySelector<HTMLButtonElement>("#generateBtn")!;
const errorsNode = document.querySelector<HTMLDivElement>("#errors")!;
const summaryNode = document.querySelector<HTMLDivElement>("#summary")!;

async function readJsonFile<T>(file: File): Promise<T> {
  const text = await file.text();
  return JSON.parse(text) as T;
}

function setErrors(lines: string[]): void {
  errorsNode.textContent = lines.join("\n");
}

function renderSummary(caseInput: CaseInput): void {
  summaryNode.hidden = false;
  summaryNode.innerHTML = `
    <strong>Summary</strong><br>
    Case Number: ${caseInput.caseNumber}<br>
    Date of Plan Termination: ${caseInput.dateOfPlanTermination}<br>
    Normal Retirement Age: ${caseInput.normalRetirementAge}<br>
    Sheets: ${PF_SHEET_ORDER.join(", ")}
  `;
}

async function loadTemplateBytes(): Promise<Uint8Array> {
  const res = await fetch("/PF_template.v0.7.13.xlsx");
  if (!res.ok) {
    throw new Error("Template file missing in web public assets: /PF_template.v0.7.13.xlsx");
  }
  return new Uint8Array(await res.arrayBuffer());
}

function downloadFile(bytes: Uint8Array, fileName: string): void {
  const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}

generateBtn.addEventListener("click", async () => {
  setErrors([]);
  summaryNode.hidden = true;

  try {
    const caseFile = caseFileInput.files?.[0];
    const planFile = planFileInput.files?.[0];

    if (!caseFile || !planFile) {
      setErrors(["Please upload both case.json and planFactors.json."]);
      return;
    }

    const caseInput = await readJsonFile<CaseInput>(caseFile);
    const planInput = await readJsonFile<PlanFactorsInput>(planFile);

    const caseValidation = validateCaseJson(caseInput);
    const planValidation = validatePlanFactorsJson(planInput);

    const errors = [
      ...caseValidation.errors.map((e) => `case.json ${e.path}: ${e.message}`),
      ...planValidation.errors.map((e) => `planFactors.json ${e.path}: ${e.message}`)
    ];

    if (errors.length > 0) {
      setErrors(errors);
      return;
    }

    const templateBytes = await loadTemplateBytes();
    const built = await buildPfWorkbook({
      caseInput,
      planFactorsInput: planInput,
      templateBytes
    });

    downloadFile(built.bytes, built.fileName);
    renderSummary(caseInput);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setErrors([message]);
  }
});
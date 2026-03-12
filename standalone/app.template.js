const PF_VERSION = "__PF_VERSION__";
const TEMPLATE_BASE64 = "__TEMPLATE_BASE64__";

const PF_SHEET_ORDER = [
  "ERF-1976",
  "ERF-1979",
  "ERF-Disability",
  "LRF-1978",
  "LRF-1982",
  "LRF-2011",
  "BFCF-1976 SLA->JS50",
  "BFCF-1979 SLA->JS50",
  "BFCF-1997 3CC->JS50",
  "BFCF-2011 3CC->JS50",
  "BFCF-2019 3CC->JS50"
];

const BFCF_SHEET_TO_KEY = {
  "BFCF-1976 SLA->JS50": "BFCF-1976",
  "BFCF-1979 SLA->JS50": "BFCF-1979",
  "BFCF-1997 3CC->JS50": "BFCF-1997",
  "BFCF-2011 3CC->JS50": "BFCF-2011",
  "BFCF-2019 3CC->JS50": "BFCF-2019"
};

const BFCF_SHEET_TO_PLAN_LABEL = {
  "BFCF-1976 SLA->JS50": "1976 Plan",
  "BFCF-1979 SLA->JS50": "1979 Plan",
  "BFCF-1997 3CC->JS50": "1997 Plan",
  "BFCF-2011 3CC->JS50": "2011 Plan",
  "BFCF-2019 3CC->JS50": "2019 Plan"
};

const app = document.getElementById("app");
app.innerHTML = `<main>
  <h1>PF Builder v${PF_VERSION} (Standalone)</h1>
  <p>Open this file directly, upload JSON inputs, and download the PF workbook.</p>
  <div class="grid">
    <label>case.json<input id="caseFile" type="file" accept=".json,application/json" /></label>
    <label>planFactors.json<input id="planFile" type="file" accept=".json,application/json" /></label>
  </div>
  <button id="generateBtn">Generate PF workbook</button>
  <pre id="errors"></pre>
  <section id="summary" hidden></section>
</main>`;

const caseFileInput = document.getElementById("caseFile");
const planFileInput = document.getElementById("planFile");
const generateBtn = document.getElementById("generateBtn");
const errorsNode = document.getElementById("errors");
const summaryNode = document.getElementById("summary");

function b64ToUint8Array(base64) {
  const bin = atob(base64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i += 1) out[i] = bin.charCodeAt(i);
  return out;
}

function formatDateMmDdYyyy(isoDate) {
  const [year, month, day] = isoDate.split("-");
  return month + "/" + day + "/" + year;
}

function fmtPercent(value, digits) {
  return (value * 100).toFixed(digits) + "%";
}

function parseJsonOrThrow(text, label) {
  try { return JSON.parse(text); }
  catch (e) { throw new Error(label + " is not valid JSON: " + e.message); }
}

function validateCaseJson(v) {
  const errors = [];
  if (!v || typeof v !== "object") errors.push("case.json must be an object");
  if (!/^\d+$/.test(v?.caseNumber ?? "")) errors.push("caseNumber must be digits only");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(v?.dateOfPlanTermination ?? "")) errors.push("dateOfPlanTermination must be yyyy-mm-dd");
  if (!Number.isInteger(v?.normalRetirementAge)) errors.push("normalRetirementAge must be an integer");
  if (typeof v?.planName !== "string" || !v.planName.trim()) errors.push("planName is required");
  return errors;
}

function validatePlanFactorsJson(v) {
  const errors = [];
  if (!v || typeof v !== "object") errors.push("planFactors.json must be an object");
  const requiredTop = ["earlyRetirement", "lateRetirement", "benefitFormConversion", "normalSingleForm"];
  for (const key of requiredTop) if (!v?.[key]) errors.push("Missing section: " + key);
  if (v?.benefitFormConversion) {
    for (const key of ["BFCF-1976","BFCF-1979","BFCF-1997","BFCF-2011","BFCF-2019"]) {
      const row = v.benefitFormConversion[key];
      if (!row) { errors.push("Missing benefitFormConversion." + key); continue; }
      if (row.toFormAbbr !== "JS50") errors.push(key + ".toFormAbbr must be JS50");
      if (row.toSurvivorPercent !== 0.5) errors.push(key + ".toSurvivorPercent must be 0.5");
      if (row.annuityType !== "MP") errors.push(key + ".annuityType must be MP");
      if (row.method !== "N") errors.push(key + ".method must be N");
    }
  }
  return errors;
}

function getSheetFileMap(workbookXml, relsXml) {
  const wbDoc = new DOMParser().parseFromString(workbookXml, "text/xml");
  const relDoc = new DOMParser().parseFromString(relsXml, "text/xml");
  const relMap = {};
  Array.from(relDoc.getElementsByTagName("Relationship")).forEach((rel) => {
    const id = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    if (id && target) relMap[id] = target.replace(/^\/?xl\//, "");
  });
  const map = {};
  Array.from(wbDoc.getElementsByTagName("sheet")).forEach((sheet) => {
    const name = sheet.getAttribute("name");
    const rid = sheet.getAttribute("r:id");
    if (name && rid && relMap[rid]) map[name] = "xl/" + relMap[rid];
  });
  return map;
}

function getRowPart(cellRef) { return Number(cellRef.match(/(\d+)$/)[1]); }
function getColPart(cellRef) { return cellRef.match(/^[A-Z]+/)[0]; }
function colToNum(col) { let n = 0; for (let i = 0; i < col.length; i += 1) n = n * 26 + (col.charCodeAt(i) - 64); return n; }
function numToCol(num) { let n = num, out = ""; while (n > 0) { const r = (n - 1) % 26; out = String.fromCharCode(65 + r) + out; n = Math.floor((n - 1) / 26); } return out; }

function iterateRange(startCol, endCol, startRow, endRow) {
  const out = [];
  const s = colToNum(startCol);
  const e = colToNum(endCol);
  for (let r = startRow; r <= endRow; r += 1) for (let c = s; c <= e; c += 1) out.push(numToCol(c) + r);
  return out;
}

function getLeftAdjacentCell(cellRef) { return numToCol(colToNum(getColPart(cellRef)) - 1) + getRowPart(cellRef); }

function ensureChildElement(doc, parent, name) {
  const existing = Array.from(parent.childNodes).find((n) => n.nodeType === 1 && n.tagName === name);
  if (existing) return existing;
  const el = doc.createElement(name);
  parent.appendChild(el);
  return el;
}

function getOrCreateCell(doc, worksheet, cellRef) {
  const sheetData = ensureChildElement(doc, worksheet, "sheetData");
  const rowNum = getRowPart(cellRef);
  let row = Array.from(sheetData.getElementsByTagName("row")).find((r) => Number(r.getAttribute("r")) === rowNum);
  if (!row) { row = doc.createElement("row"); row.setAttribute("r", String(rowNum)); sheetData.appendChild(row); }
  let cell = Array.from(row.getElementsByTagName("c")).find((c) => c.getAttribute("r") === cellRef);
  if (!cell) { cell = doc.createElement("c"); cell.setAttribute("r", cellRef); row.appendChild(cell); }
  return cell;
}

function clearCellChildren(cell) { while (cell.firstChild) cell.removeChild(cell.firstChild); }

function setInlineStringCell(doc, worksheet, cellRef, value) {
  const cell = getOrCreateCell(doc, worksheet, cellRef);
  cell.setAttribute("t", "inlineStr");
  clearCellChildren(cell);
  const isNode = doc.createElement("is");
  const tNode = doc.createElement("t");
  tNode.textContent = value;
  isNode.appendChild(tNode);
  cell.appendChild(isNode);
}

function setFormulaCell(doc, worksheet, cellRef, formula) {
  const cell = getOrCreateCell(doc, worksheet, cellRef);
  cell.removeAttribute("t");
  clearCellChildren(cell);
  const f = doc.createElement("f");
  f.textContent = formula.replace(/^=/, "");
  cell.appendChild(f);
  const v = doc.createElement("v");
  v.textContent = "0";
  cell.appendChild(v);
}

function quote(value) { return '"' + value + '"'; }

function erfFormula(sheetName, monthsCell, nra, plan) {
  if (sheetName === "ERF-1976") return `=ROUND(1-(${monthsCell}/180),4)`;
  if (sheetName === "ERF-1979") return `=ROUND(1-(1/180)*MIN(${monthsCell},60)-(1/360)*MAX(MIN(${monthsCell}-60,60),0),4)`;
  const d = plan.earlyRetirement["ERF-Disability"];
  return `=ROUND(ERFAEQ(${d.interest},${quote("M")},${nra},${monthsCell},${d.monthsCertain},${quote(d.mortalityMale)},${quote(d.mortalityFemale)},${quote("MP")},${quote(d.deferralMortality)}),4)`;
}

function lrfFormula(monthsCell, a, nra) {
  return `=ROUND(LRFAEQ(${a.interest},${quote("M")},${nra},${monthsCell},${a.monthsCertain},${quote(a.mortalityMale)},${quote(a.mortalityFemale)},${quote("MP")},${quote(a.deferralMortality)}),4)`;
}

function bfcfFormula(cellRef, a) {
  const participantAgeCell = cellRef.replace(/[0-9]+$/, "") + "10";
  const beneficiaryAgeCell = "B" + cellRef.match(/[0-9]+$/)[0];
  return `=ROUND(BFCFAEQ(${quote(a.fromFormAbbr)},0,${a.fromMonthsCertain},0,${quote("JSC")},0.5,0,0,${participantAgeCell},${beneficiaryAgeCell},${quote("M")},${quote("M")},${a.interest},${quote(a.mortalityMale)},${quote(a.mortalityFemale)},${quote("MP")},${quote("N")}),4)`;
}

function normalSingleFormForSheet(sheetName, n) {
  if (sheetName === "ERF-1976") return n["1976"];
  if (sheetName === "ERF-1979") return n["1979"];
  if (sheetName === "ERF-Disability") return n["2011"];
  if (sheetName === "LRF-1978") return n["1976"];
  if (sheetName === "LRF-1982") return n["1997"];
  if (sheetName === "LRF-2011") return n["2011"];
  if (sheetName === "BFCF-1976 SLA->JS50") return n["1976"];
  if (sheetName === "BFCF-1979 SLA->JS50") return n["1979"];
  if (sheetName === "BFCF-1997 3CC->JS50") return n["1997"];
  if (sheetName === "BFCF-2011 3CC->JS50") return n["2011"];
  return n["2019"];
}

function setCommonHeader(caseInput, doc, worksheet, baseCol) {
  setInlineStringCell(doc, worksheet, baseCol + "1", caseInput.planName);
  setInlineStringCell(doc, worksheet, baseCol + "2", "Case Number: " + caseInput.caseNumber);
  setInlineStringCell(doc, worksheet, baseCol + "3", "DOPT: " + formatDateMmDdYyyy(caseInput.dateOfPlanTermination));
}

function patchErfSheet(sheetName, caseInput, plan, doc, ws) {
  setCommonHeader(caseInput, doc, ws, "A");
  if (sheetName === "ERF-1976") {
    setInlineStringCell(doc, ws, "A5", `Early Retirement Factors from NRA of ${caseInput.normalRetirementAge} (1976 Plan)`);
    setInlineStringCell(doc, ws, "A6", "Basis: reduction = 1/15 per year early (linear in months)");
  } else if (sheetName === "ERF-1979") {
    setInlineStringCell(doc, ws, "A5", `Early Retirement Factors from NRA of ${caseInput.normalRetirementAge} (1979/1997/2011/2019 Plans)`);
    setInlineStringCell(doc, ws, "A6", "Basis: reduction = 1/180 per month for first 60 months, then 1/360 per month for next 60 months");
  } else {
    const d = plan.earlyRetirement["ERF-Disability"];
    setInlineStringCell(doc, ws, "A5", `Early Retirement Factors from NRA of ${caseInput.normalRetirementAge} (Disability (2011/2019))`);
    setInlineStringCell(doc, ws, "A6", `Basis: ERFAEQ with interest=${fmtPercent(d.interest,0)}; mortality=${d.mortalityMale}/${d.mortalityFemale}; calc method=MP; deferral mortality=${d.deferralMortality}; months certain=${d.monthsCertain}`);
  }
  setInlineStringCell(doc, ws, "A7", `Normal single form: ${normalSingleFormForSheet(sheetName, plan.normalSingleForm)}`);
  for (let r = 11; r <= 34; r += 1) {
    for (const c of ["B","D","F","H","J"]) {
      const target = c + r;
      setFormulaCell(doc, ws, target, erfFormula(sheetName, getLeftAdjacentCell(target), String(caseInput.normalRetirementAge), plan));
    }
  }
}

function patchLrfSheet(sheetName, caseInput, plan, doc, ws) {
  setCommonHeader(caseInput, doc, ws, "A");
  const a = plan.lateRetirement[sheetName];
  if (sheetName === "LRF-1978") setInlineStringCell(doc, ws, "A5", `Late Retirement Factors from NRA of ${caseInput.normalRetirementAge} (1978 Plan)`);
  else if (sheetName === "LRF-1982") setInlineStringCell(doc, ws, "A5", `Late Retirement Factors from NRA of ${caseInput.normalRetirementAge} (1982 Plan)`);
  else setInlineStringCell(doc, ws, "A5", `Late Retirement Factors from NRA of ${caseInput.normalRetirementAge} (2011/2019 Plans)`);
  setInlineStringCell(doc, ws, "A6", `Basis: LRFAEQ with interest=${fmtPercent(a.interest,0)}; mortality=${a.mortalityMale}/${a.mortalityFemale}; calc method=MP; deferral mortality=${a.deferralMortality}; months certain=${a.monthsCertain}`);
  setInlineStringCell(doc, ws, "A7", `Normal single form: ${normalSingleFormForSheet(sheetName, plan.normalSingleForm)}`);
  for (let r = 11; r <= 34; r += 1) {
    for (const c of ["B","D","F","H","J"]) {
      const target = c + r;
      setFormulaCell(doc, ws, target, lrfFormula(getLeftAdjacentCell(target), a, String(caseInput.normalRetirementAge)));
    }
  }
}

function patchBfcfSheet(sheetName, caseInput, plan, doc, ws) {
  setCommonHeader(caseInput, doc, ws, "B");
  const key = BFCF_SHEET_TO_KEY[sheetName];
  const a = plan.benefitFormConversion[key];
  setInlineStringCell(doc, ws, "B5", `Form Conversion Factors: ${a.fromFormAbbr}->JS50 (${BFCF_SHEET_TO_PLAN_LABEL[sheetName]})`);
  setInlineStringCell(doc, ws, "B6", `Basis: interest=${fmtPercent(a.interest,2)}; mortality=${a.mortalityMale}/${a.mortalityFemale}; method=MP`);
  setInlineStringCell(doc, ws, "B7", `Normal single form: ${normalSingleFormForSheet(sheetName, plan.normalSingleForm)}`);
  for (const cellRef of iterateRange("C","AB",11,111)) setFormulaCell(doc, ws, cellRef, bfcfFormula(cellRef, a));
}

async function buildWorkbook(caseInput, planInput) {
  const zip = await JSZip.loadAsync(b64ToUint8Array(TEMPLATE_BASE64));
  const workbookXml = await zip.file("xl/workbook.xml").async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const sheetMap = getSheetFileMap(workbookXml, relsXml);

  for (const sheetName of PF_SHEET_ORDER) {
    const path = sheetMap[sheetName];
    if (!path) throw new Error("Missing sheet path: " + sheetName);
    const xml = await zip.file(path).async("string");
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const ws = doc.getElementsByTagName("worksheet")[0];

    if (sheetName.startsWith("ERF")) patchErfSheet(sheetName, caseInput, planInput, doc, ws);
    else if (sheetName.startsWith("LRF")) patchLrfSheet(sheetName, caseInput, planInput, doc, ws);
    else patchBfcfSheet(sheetName, caseInput, planInput, doc, ws);

    zip.file(path, new XMLSerializer().serializeToString(doc));
  }

  const out = await zip.generateAsync({ type: "uint8array" });
  return { bytes: out, fileName: `${caseInput.caseNumber}PF.v${PF_VERSION}.xlsx` };
}

function setErrors(lines) { errorsNode.textContent = lines.join("\n"); }

function downloadXlsx(bytes, fileName) {
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
      setErrors(["Please upload case.json and planFactors.json."]);
      return;
    }

    const caseInput = parseJsonOrThrow(await caseFile.text(), "case.json");
    const planInput = parseJsonOrThrow(await planFile.text(), "planFactors.json");

    const errors = [...validateCaseJson(caseInput), ...validatePlanFactorsJson(planInput)];
    if (errors.length) {
      setErrors(errors);
      return;
    }

    const built = await buildWorkbook(caseInput, planInput);
    downloadXlsx(built.bytes, built.fileName);

    summaryNode.hidden = false;
    summaryNode.innerHTML =
      "<h3>Summary</h3>" +
      "<div>Case: " + caseInput.caseNumber + "</div>" +
      "<div>DOPT: " + caseInput.dateOfPlanTermination + "</div>" +
      "<div>NRA: " + caseInput.normalRetirementAge + "</div>" +
      "<div>Sheets: " + PF_SHEET_ORDER.join(", ") + "</div>";
  } catch (err) {
    setErrors([err instanceof Error ? err.message : String(err)]);
  }
});

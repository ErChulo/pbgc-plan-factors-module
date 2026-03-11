# Module PF Builder Prompt for SpecKit + Codex CLI (PF v0.7.13)

You are Codex CLI working under SpecKit. Build a **Plan Factors (PF) workbook builder** module that generates an Excel workbook **identical in layout, formatting, print settings, merged cells, formulas, and sheet names** to the reference workbook:

- Reference: `24884900PF.v0.7.13.xlsx`

The module must accept plan data and case data as JSON inputs (uploaded by the user in the web application) and output:

- `<caseNumber>PF.v0.7.13.xlsx`

The generated file must match the reference workbook’s **structure and formatting** exactly (same sheet list, same cell merges, same borders, same print titles, same print area logic, same freeze panes, same column widths). Only the **content values that come from inputs** (for example, case number, date of plan termination, actuarial assumptions, and formula parameters) are allowed to differ when inputs differ.

---

## 0. Non-negotiable constraints

1) **Exact formatting parity with PF v0.7.13**
- Do **not** “recreate” the workbook from scratch.
- Use a **binary template** file stored in the repository:  
  `templates/PF_template.v0.7.13.xlsx`  
  (This is a copy of `24884900PF.v0.7.13.xlsx`.)

2) **Patch, don’t rebuild**
- The builder must **copy** the template and then **patch only specific cells** (values + formulas) based on inputs.
- All other workbook parts must remain byte-for-byte stable whenever inputs are the same.

3) **Client-side web application**
- The graphical user interface must run fully client-side (no server).
- The user uploads JSON files; the app generates the workbook and triggers a download.

4) **Versioning**
- Hardcode PF builder output version for this module: **v0.7.13**.
- Store it in a single constant `PF_VERSION = "0.7.13"`.

---

## 1. Repository deliverables

### 1.1 Packages

- `packages/core-pf`
  - Pure library that:
    - validates inputs against JSON schema,
    - patches the template workbook,
    - returns the final `.xlsx` bytes.

- `packages/web-pf`
  - Vite + TypeScript web application.
  - User uploads JSON inputs and downloads the generated PF workbook.

- `packages/schemas`
  - JSON schema files for input validation.

- `templates/`
  - `PF_template.v0.7.13.xlsx` (binary, committed to repository)

- `fixtures/`
  - `case.sample.json`
  - `planFactors.sample.json`
  - `expected.24884900PF.v0.7.13.xlsx` (golden file fixture for tests)

### 1.2 Command line tool (optional but strongly recommended)
- `packages/cli-pf`
  - `pf-build --case case.json --plan planFactors.json --out outdir/`
  - Must generate the same output as the web application for the same inputs.

---

## 2. Inputs

The user uploads **two JSON files**.

### 2.1 Case JSON (required)

Filename: `case.json`

```json
{
  "caseNumber": "24884900",
  "planName": "The College of Saint Rose Non-Contract Employees Pension Plan",
  "dateOfPlanTermination": "2024-06-30",
  "normalRetirementAge": 65
}
```

Rules:
- `caseNumber` required; must be digits only.
- `dateOfPlanTermination` required; ISO `yyyy-mm-dd`.
- `normalRetirementAge` required; integer.
- The PF builder prints DOPT in `mm/dd/yyyy` in the workbook heading.

### 2.2 Plan factors JSON (required)

Filename: `planFactors.json`

This file describes the parameters used to fill the formulas for each factor family and each document-year tab.

Required sections:

```json
{
  "earlyRetirement": {
    "ERF-1976": { "type": "linear_1_over_15_per_year" },
    "ERF-1979": { "type": "tiered_monthly_1_180_then_1_360" },
    "ERF-Disability": {
      "type": "actuarial",
      "interest": 0.08,
      "sex": "M",
      "monthsCertain": 36,
      "mortalityMale": "UP84M2",
      "mortalityFemale": "UP84M2",
      "calcMethod": "MP",
      "deferralMortality": "Y"
    }
  },
  "lateRetirement": {
    "LRF-1978": {
      "interest": 0.05,
      "sex": "M",
      "monthsCertain": 0,
      "mortalityMale": "GM83",
      "mortalityFemale": "GF83",
      "calcMethod": "MP",
      "deferralMortality": "Y"
    },
    "LRF-1982": {
      "interest": 0.05,
      "sex": "M",
      "monthsCertain": 36,
      "mortalityMale": "UP84",
      "mortalityFemale": "UP84M2",
      "calcMethod": "MP",
      "deferralMortality": "Y"
    },
    "LRF-2011": {
      "interest": 0.08,
      "sex": "M",
      "monthsCertain": 36,
      "mortalityMale": "UP84M2",
      "mortalityFemale": "UP84M2",
      "calcMethod": "MP",
      "deferralMortality": "Y"
    }
  },
  "benefitFormConversion": {
    "BFCF-1976": {
      "fromFormAbbr": "SLA",
      "fromMonthsCertain": 0,
      "toFormAbbr": "JS50",
      "toSurvivorPercent": 0.5,
      "interest": 0.05,
      "mortalityMale": "GM83",
      "mortalityFemale": "GM83",
      "annuityType": "MP",
      "method": "N"
    },
    "BFCF-1979": {
      "fromFormAbbr": "SLA",
      "fromMonthsCertain": 0,
      "toFormAbbr": "JS50",
      "toSurvivorPercent": 0.5,
      "interest": 0.05,
      "mortalityMale": "GM83",
      "mortalityFemale": "GM83",
      "annuityType": "MP",
      "method": "N"
    },
    "BFCF-1997": {
      "fromFormAbbr": "3CC",
      "fromMonthsCertain": 36,
      "toFormAbbr": "JS50",
      "toSurvivorPercent": 0.5,
      "interest": 0.08,
      "mortalityMale": "UP84",
      "mortalityFemale": "UP84",
      "annuityType": "MP",
      "method": "N"
    },
    "BFCF-2011": {
      "fromFormAbbr": "3CC",
      "fromMonthsCertain": 36,
      "toFormAbbr": "JS50",
      "toSurvivorPercent": 0.5,
      "interest": 0.08,
      "mortalityMale": "UP84M2",
      "mortalityFemale": "UP84M2",
      "annuityType": "MP",
      "method": "N"
    },
    "BFCF-2019": {
      "fromFormAbbr": "3CC",
      "fromMonthsCertain": 36,
      "toFormAbbr": "JS50",
      "toSurvivorPercent": 0.5,
      "interest": 0.08,
      "mortalityMale": "UP84M2",
      "mortalityFemale": "UP84M2",
      "annuityType": "MP",
      "method": "N"
    }
  },
  "normalSingleForm": {
    "1976": "SLA",
    "1979": "SLA",
    "1997": "3CC",
    "2011": "3CC",
    "2019": "3CC"
  }
}
```

Rules:
- All interest values are decimals (example: 0.08 for 8%).
- Mortality handles are strings (passed into add-in functions unchanged).
- `calcMethod` must be `"MP"` for late retirement and disability tables in this module.
- `fromFormAbbr` limited to: `SLA`, `3CC`.
- `toFormAbbr` is `JS50` only for this module.

---

## 3. JSON schema validation

Create JSON schema files in `packages/schemas`:

- `case.schema.json`
- `planFactors.schema.json`

Validation requirements:
- Validate user-uploaded JSON before generating.
- Show validation errors in the web application (human-readable list).
- Do not generate output if invalid.

---

## 4. Workbook specification (must match PF v0.7.13)

### 4.1 Sheet list (exact order)

1. `ERF-1976`
2. `ERF-1979`
3. `ERF-Disability`
4. `LRF-1978`
5. `LRF-1982`
6. `LRF-2011`
7. `BFCF-1976 SLA->JS50`
8. `BFCF-1979 SLA->JS50`
9. `BFCF-1997 3CC->JS50`
10. `BFCF-2011 3CC->JS50`
11. `BFCF-2019 3CC->JS50`

### 4.2 Print and freeze settings (must match)

For **ERF and LRF sheets**:
- Freeze panes: `A11`
- Print title rows: `1:10`
- Print area: `A1:J34`
- Orientation: landscape
- Fit to: **1 page wide by 2 pages tall**
- Left margin: 0.4

For **BFCF sheets**:
- Freeze panes: `C11`
- Print title rows: `1:10`
- Print title columns: `A:B`
- Print area: `A1:AB111`
- Orientation: landscape
- Fit to: **1 page wide by 2 pages tall**
- Left margin: 0.4

### 4.3 BFCF age ranges (must match)

- Participant age nearest birthday: **55 to 80** (columns `C`..`AB`).
- Beneficiary age nearest birthday: **0 to 100** (rows `11`..`111`).

### 4.4 Column widths (BFCF)

- Column A width: 4.0
- Column B width: 8.5
- Columns C..AB width: 8.5 (must avoid `####` rendering in normal view)

### 4.5 Labels and merged cells (BFCF)

- Column A has two merged label blocks:
  - `A11:A70` labeled “Beneficiary's Age Nearest Birthday”, rotated 90 degrees, vertical top, centered horizontal, bold, bordered.
  - `A71:A111` same label and formatting and borders.
- Row 9: `C9:AB9` merged; value “Participant's Age Nearest Birthday”, centered, bold, bordered.

### 4.6 Factor formatting

- All factor cells must have Excel number format: `0.0000`
- Formulas must wrap with `ROUND(…,4)`.

---

## 5. Formula rules (patch targets)

### 5.1 ERF factor cells
In each ERF sheet, factor cells are the even columns in the 5 month blocks (B, D, F, H, J) for rows 11..34.

- `ERF-1976`:
  - `=ROUND(1-(MonthsCell/180),4)`

- `ERF-1979`:
  - `=ROUND(1-(1/180)*MIN(MonthsCell,60)-(1/360)*MAX(MIN(MonthsCell-60,60),0),4)`

- `ERF-Disability`:
  - `=ROUND(ERFAEQ(interest,"M",NRA,MonthsCell,monthsCertain,"mortM","mortF","MP","Y"),4)`
  - `interest`, `monthsCertain`, `mortM`, `mortF` come from planFactors JSON.

### 5.2 LRF factor cells
In each LRF sheet, factor cells are the even columns in the 5 month blocks (B, D, F, H, J) for rows 11..34.

- General:
  - `=ROUND(LRFAEQ(interest,"M",NRA,MonthsCell,monthsCertain,"mortM","mortF","MP","Y"),4)`

### 5.3 BFCF matrix cells
In each BFCF sheet, the factor grid is `C11:AB111`.

Use `BFCFAEQ` (not `PBGCBFCF2`) and pass the plan assumptions:

```
=ROUND(
  BFCFAEQ(
    "SLA",0,fromMonthsCertain,0,
    "JSC",0.5,0,0,
    ParticipantAgeCell,BeneficiaryAgeCell,
    "M","M",
    interest,
    "mortM","mortF",
    "MP","N"
  ),
4)
```

Notes:
- The **from form** is `SLA` with months certain 0 or 36 depending on tab.
- The **to form** is `JSC` with survivor percent 0.5 (JS50).
- `interest`, `mortM`, `mortF` come from planFactors JSON.

---

## 6. Implementation approach (required)

### 6.1 Use OOXML patching to preserve formatting
Implement the patcher using:
- `JSZip` to load the `.xlsx` (zip).
- XML parsing (DOMParser) to modify:
  - worksheet XML for specific cells (`v` and `f` nodes),
  - shared strings (only if you choose to store heading text as shared strings; simplest is inline strings for patched cells).

Do **not** use a high-level Excel writer that might rewrite styles and break parity.

### 6.2 Patch by coordinates
Maintain a mapping file in code that lists:
- sheet name
- cell address or range
- patch operation (set string, set formula, set number format if needed)

---

## 7. Web application requirements

User interface:
- Two upload widgets: `case.json` and `planFactors.json`.
- Validate both (schema validation). Display errors.
- A button: “Generate PF workbook”.
- Download prompt for `<caseNumber>PF.v0.7.13.xlsx`.
- Show a summary card of:
  - case number
  - date of plan termination
  - normal retirement age
  - which sheets will be generated

No server. No external network calls.

---

## 8. Automated tests (required)

### 8.1 Golden-file parity test
Using the fixtures:
- Load `case.sample.json` and `planFactors.sample.json`
- Generate workbook bytes
- Compare SHA256 hash to `fixtures/expected.24884900PF.v0.7.13.xlsx`

### 8.2 Structural tests
Programmatically inspect generated workbook:
- sheet order matches expected list
- print area and print titles match (per sheet)
- freeze panes match (per sheet)
- merged ranges match for BFCF sheets
- column widths match for BFCF sheets

### 8.3 Formula tests
Spot-check a handful of cells:
- ERF factor cell formula contains `ROUND` and correct structure
- LRF factor cell uses `LRFAEQ` and `MP`
- BFCF cell uses `BFCFAEQ` and includes interest and mortality

---

## 9. Output naming

Output filename:
- `${caseNumber}PF.v0.7.13.xlsx`

---

## 10. Start now

Implement **exactly** what is described above. Prioritize parity with the template and the golden-file test above all else.

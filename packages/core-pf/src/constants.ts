export const PF_VERSION = "0.7.13" as const;

export const PF_SHEET_ORDER = [
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
] as const;

export type PfSheetName = (typeof PF_SHEET_ORDER)[number];

export const HEADER_CELL_MAP = {
  caseNumber: "B2",
  planName: "B3",
  dopt: "H2",
  nra: "H3"
} as const;

export const FACTOR_STYLE_HINT = "0.0000";
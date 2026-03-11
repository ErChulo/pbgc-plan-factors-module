import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

const worksheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, "text/xml");
}

function ensureChildElement(doc: Document, parent: Element, tagName: string): Element {
  const existing = Array.from(parent.childNodes).find(
    (n) => n.nodeType === 1 && (n as Element).tagName === tagName
  ) as Element | undefined;

  if (existing) {
    return existing;
  }

  const child = doc.createElementNS(worksheetNs, tagName);
  parent.appendChild(child);
  return child;
}

function getColumnPart(cellRef: string): string {
  const match = /^[A-Z]+/.exec(cellRef);
  if (!match) throw new Error(`Invalid cell reference: ${cellRef}`);
  return match[0];
}

function getRowPart(cellRef: string): number {
  const match = /([0-9]+)$/.exec(cellRef);
  if (!match) throw new Error(`Invalid cell reference: ${cellRef}`);
  return Number(match[1]);
}

export function formatDateMmDdYyyy(isoDate: string): string {
  const [year, month, day] = isoDate.split("-");
  return `${month}/${day}/${year}`;
}

function findRow(sheetData: Element, rowNumber: number): Element | undefined {
  const rows = Array.from(sheetData.getElementsByTagName("row"));
  return rows.find((row) => Number(row.getAttribute("r")) === rowNumber);
}

function createRow(doc: Document, sheetData: Element, rowNumber: number): Element {
  const row = doc.createElementNS(worksheetNs, "row");
  row.setAttribute("r", String(rowNumber));
  sheetData.appendChild(row);
  return row;
}

function findCell(row: Element, cellRef: string): Element | undefined {
  const cells = Array.from(row.getElementsByTagName("c"));
  return cells.find((cell) => cell.getAttribute("r") === cellRef);
}

function createCell(doc: Document, row: Element, cellRef: string): Element {
  const cell = doc.createElementNS(worksheetNs, "c");
  cell.setAttribute("r", cellRef);
  row.appendChild(cell);
  return cell;
}

function getOrCreateCell(doc: Document, worksheet: Element, cellRef: string): Element {
  const sheetData = ensureChildElement(doc, worksheet, "sheetData");
  const rowNumber = getRowPart(cellRef);
  const row = findRow(sheetData, rowNumber) ?? createRow(doc, sheetData, rowNumber);
  return findCell(row, cellRef) ?? createCell(doc, row, cellRef);
}

function clearCellChildren(cell: Element): void {
  while (cell.firstChild) {
    cell.removeChild(cell.firstChild);
  }
}

export function setFormulaCell(doc: Document, worksheet: Element, cellRef: string, formula: string): void {
  const cell = getOrCreateCell(doc, worksheet, cellRef);
  cell.removeAttribute("t");
  clearCellChildren(cell);

  const f = doc.createElementNS(worksheetNs, "f");
  f.appendChild(doc.createTextNode(formula.replace(/^=/, "")));
  cell.appendChild(f);

  const v = doc.createElementNS(worksheetNs, "v");
  v.appendChild(doc.createTextNode("0"));
  cell.appendChild(v);
}

export function setInlineStringCell(doc: Document, worksheet: Element, cellRef: string, value: string): void {
  const cell = getOrCreateCell(doc, worksheet, cellRef);
  cell.setAttribute("t", "inlineStr");
  clearCellChildren(cell);

  const isNode = doc.createElementNS(worksheetNs, "is");
  const tNode = doc.createElementNS(worksheetNs, "t");
  tNode.appendChild(doc.createTextNode(value));
  isNode.appendChild(tNode);
  cell.appendChild(isNode);
}

export function parseWorksheetXml(xml: string): { doc: Document; worksheet: Element } {
  const doc = parseXml(xml);
  const worksheet = doc.getElementsByTagName("worksheet")[0];
  if (!worksheet) {
    throw new Error("Invalid worksheet XML");
  }
  return { doc, worksheet };
}

export function serializeXml(doc: Document): string {
  return new XMLSerializer().serializeToString(doc);
}

export function columnToIndex(column: string): number {
  let n = 0;
  for (let i = 0; i < column.length; i += 1) {
    n = n * 26 + (column.charCodeAt(i) - 64);
  }
  return n;
}

export function indexToColumn(index: number): string {
  let value = index;
  let out = "";
  while (value > 0) {
    const rem = (value - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    value = Math.floor((value - 1) / 26);
  }
  return out;
}

export function iterateRange(startCol: string, endCol: string, startRow: number, endRow: number): string[] {
  const cells: string[] = [];
  const start = columnToIndex(startCol);
  const end = columnToIndex(endCol);

  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = start; col <= end; col += 1) {
      cells.push(`${indexToColumn(col)}${row}`);
    }
  }
  return cells;
}

export function getLeftAdjacentCell(cellRef: string): string {
  const col = getColumnPart(cellRef);
  const row = getRowPart(cellRef);
  return `${indexToColumn(columnToIndex(col) - 1)}${row}`;
}
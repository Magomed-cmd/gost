import ExcelJS from "exceljs";

export interface ExcelStyleConfig {
  HEADER_FILL: string;
  RESULT_FILL: string;
  ALT_ROW_FILL: string;
  REGRET_FILL: string;
  FONT_NAME: string;
  FONT_SIZE: number;
  TAB_COLOR: string;
  DEFAULT_ROW_HEIGHT: number;
  PAGE_ORIENTATION: "landscape" | "portrait";
  FIT_TO_PAGE: boolean;
  FIT_TO_WIDTH: number;
  FIT_TO_HEIGHT: number;
  MAX_COLUMN_WIDTH: number;
  MAX_ROW_HEIGHT: number;
  ROW_HEIGHT_FONT_MULTIPLIER: number;
  HEADER_TITLE_SIZE: number;
  BORDER_STYLE: string;
}

export interface CellStyleOptions {
  fontName?: string;
  size?: number;
  bold?: boolean;
  color?: string;
  align?: "left" | "center" | "right" | "fill" | "justify" | "centerContinuous" | "distributed";
  wrapText?: boolean;
  fill?: string;
  border?: boolean | string;
  numFmt?: string;
}

export interface AhpBlockResult {
  weightColumn: string;
  weightStartRow: number;
  weightEndRow: number;
  osCell: string;
  nextRow: number;
}

export interface ExcelGostInstance {
  readonly style: ExcelStyleConfig;
  readonly HEADER_STYLE: CellStyleOptions;
  readonly RESULT_STYLE: CellStyleOptions;
  styleCell(cell: ExcelJS.Cell, options?: CellStyleOptions): void;
  setValue(sheet: ExcelJS.Worksheet, address: string, value: unknown, options?: CellStyleOptions): ExcelJS.Cell;
  setFormula(sheet: ExcelJS.Worksheet, address: string, formula: string, options?: CellStyleOptions): ExcelJS.Cell;
  applyBaseSheetSetup(sheet: ExcelJS.Worksheet, widths: number[]): void;
  mergeTitle(sheet: ExcelJS.Worksheet, range: string, text: string): void;
  setFreeze(sheet: ExcelJS.Worksheet, ySplit: number, activeCell?: string): void;
  applyZebraStriping(sheet: ExcelJS.Worksheet, startRow: number, endRow: number, startCol: number, endCol: number, headerRows?: Set<number>): void;
  getCellText(cell: ExcelJS.Cell): string;
  autoFitSheet(sheet: ExcelJS.Worksheet): void;
  formatTableRange(sheet: ExcelJS.Worksheet, startRow: number, endRow: number, startCol: number, endCol: number, headerRow: number): void;
  addAhpMatrixBlock(sheet: ExcelJS.Worksheet, startRow: number, title: string, labels: string[], matrix: number[][], randomIndex: number): AhpBlockResult;
  addContentsSheet(workbook: ExcelJS.Workbook, sheetNames: string[]): ExcelJS.Worksheet;
  productFormula(addresses: string[], exponent: number): string;
  absoluteRef(address: string): string;
  weightedSumFormula(valueAddresses: string[], weightAddresses: string[]): string;
  listFormulaRange(sheetName: string, addresses: string[]): string;
}

export function createExcelStyle(opts: Partial<ExcelStyleConfig> = {}): ExcelStyleConfig {
  return {
    HEADER_FILL: opts.HEADER_FILL ?? "FFF2CC",
    RESULT_FILL: opts.RESULT_FILL ?? "D9EAD3",
    ALT_ROW_FILL: opts.ALT_ROW_FILL ?? "FFF9F0",
    REGRET_FILL: opts.REGRET_FILL ?? "FCE4D6",
    FONT_NAME: opts.FONT_NAME ?? "Times New Roman",
    FONT_SIZE: opts.FONT_SIZE ?? 14,
    TAB_COLOR: opts.TAB_COLOR ?? "FFF2CC",
    DEFAULT_ROW_HEIGHT: opts.DEFAULT_ROW_HEIGHT ?? 20,
    PAGE_ORIENTATION: opts.PAGE_ORIENTATION ?? "landscape",
    FIT_TO_PAGE: opts.FIT_TO_PAGE ?? true,
    FIT_TO_WIDTH: opts.FIT_TO_WIDTH ?? 1,
    FIT_TO_HEIGHT: opts.FIT_TO_HEIGHT ?? 0,
    MAX_COLUMN_WIDTH: opts.MAX_COLUMN_WIDTH ?? 80,
    MAX_ROW_HEIGHT: opts.MAX_ROW_HEIGHT ?? 120,
    ROW_HEIGHT_FONT_MULTIPLIER: opts.ROW_HEIGHT_FONT_MULTIPLIER ?? 1.9,
    HEADER_TITLE_SIZE: opts.HEADER_TITLE_SIZE ?? 14,
    BORDER_STYLE: opts.BORDER_STYLE ?? "thin",
  };
}

function styleCellImpl(cell: ExcelJS.Cell, options: CellStyleOptions = {}, st: ExcelStyleConfig): void {
  cell.font = {
    name: options.fontName ?? st.FONT_NAME,
    size: options.size ?? st.FONT_SIZE,
    bold: Boolean(options.bold),
    color: options.color ? { argb: options.color } : undefined,
  };
  cell.alignment = {
    vertical: "middle",
    horizontal: options.align ?? "center",
    wrapText: options.wrapText !== false,
  };
  if (options.fill) {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: options.fill } };
  }
  if (options.border) {
    const borderStyle = typeof options.border === "string" ? options.border : st.BORDER_STYLE;
    cell.border = {
      top: { style: borderStyle },
      left: { style: borderStyle },
      bottom: { style: borderStyle },
      right: { style: borderStyle },
    } as Partial<ExcelJS.Borders>;
  }
  if (options.numFmt) {
    cell.numFmt = options.numFmt;
  }
}

function setValueImpl(sheet: ExcelJS.Worksheet, address: string, value: unknown, options: CellStyleOptions, st: ExcelStyleConfig): ExcelJS.Cell {
  const cell = sheet.getCell(address);
  cell.value = value as ExcelJS.CellValue;
  styleCellImpl(cell, options, st);
  return cell;
}

function setFormulaImpl(sheet: ExcelJS.Worksheet, address: string, formula: string, options: CellStyleOptions, st: ExcelStyleConfig): ExcelJS.Cell {
  const cell = sheet.getCell(address);
  cell.value = { formula };
  styleCellImpl(cell, options, st);
  return cell;
}

function applyBaseSheetSetupImpl(sheet: ExcelJS.Worksheet, widths: number[], st: ExcelStyleConfig): void {
  sheet.properties.defaultRowHeight = st.DEFAULT_ROW_HEIGHT;
  sheet.pageSetup.orientation = st.PAGE_ORIENTATION;
  sheet.pageSetup.fitToPage = st.FIT_TO_PAGE;
  sheet.pageSetup.fitToWidth = st.FIT_TO_WIDTH;
  sheet.pageSetup.fitToHeight = st.FIT_TO_HEIGHT;
  (sheet as ExcelJS.Worksheet & { _baseWidths?: number[] })._baseWidths = widths;
  widths.forEach((width, index) => {
    sheet.getColumn(index + 1).width = width;
  });
}

function mergeTitleImpl(sheet: ExcelJS.Worksheet, range: string, text: string, st: ExcelStyleConfig): void {
  sheet.mergeCells(range);
  const cell = sheet.getCell(range.split(":")[0]);
  cell.value = text;
  styleCellImpl(cell, { bold: true, size: st.HEADER_TITLE_SIZE, align: "center" }, st);
}

function setFreezeImpl(sheet: ExcelJS.Worksheet, ySplit: number, activeCell = "A2"): void {
  void sheet;
  void ySplit;
  void activeCell;
}

function applyZebraStripingImpl(sheet: ExcelJS.Worksheet, startRow: number, endRow: number, startCol: number, endCol: number, headerRows = new Set<number>(), st: ExcelStyleConfig): void {
  for (let row = startRow; row <= endRow; row += 1) {
    if (headerRows.has(row) || row % 2 !== 0) continue;
    for (let col = startCol; col <= endCol; col += 1) {
      const cell = sheet.getCell(row, col);
      const fill = cell.fill as Partial<ExcelJS.FillPattern> | undefined;
      const existingFill = fill?.fgColor?.argb;
      if (existingFill === st.HEADER_FILL || existingFill === st.RESULT_FILL) continue;
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: st.ALT_ROW_FILL } };
    }
  }
}

function getCellTextImpl(cell: ExcelJS.Cell): string {
  if (!cell || cell.value === null || cell.value === undefined) return "";
  if (typeof cell.value === "object" && "formula" in cell.value) return "";
  if (typeof cell.value === "object" && "text" in cell.value) return String(cell.value.text);
  return String(cell.text || cell.value);
}

function autoFitSheetImpl(sheet: ExcelJS.Worksheet, st: ExcelStyleConfig): void {
  const baseWidths = (sheet as ExcelJS.Worksheet & { _baseWidths?: number[] })._baseWidths ?? [];
  const maxColumn = Math.max(sheet.columnCount, baseWidths.length);
  const computedWidths = Array.from({ length: maxColumn }, (_, index) => baseWidths[index] || 10);

  for (let columnIndex = 1; columnIndex <= maxColumn; columnIndex += 1) {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      const text = getCellTextImpl(row.getCell(columnIndex)).trim();
      if (!text) return;
      const longest = text.split(/\r?\n/).reduce((max, line) => Math.max(max, line.length), 0);
      computedWidths[columnIndex - 1] = Math.min(st.MAX_COLUMN_WIDTH, Math.max(computedWidths[columnIndex - 1], longest + 2));
    });
  }
  computedWidths.forEach((width, index) => { sheet.getColumn(index + 1).width = width; });

  sheet.eachRow({ includeEmpty: false }, (row) => {
    let rowHeight = sheet.properties.defaultRowHeight || st.DEFAULT_ROW_HEIGHT;
    row.eachCell({ includeEmpty: false }, (cell) => {
      const text = getCellTextImpl(cell).trim();
      if (!text) return;
      const width = sheet.getColumn(cell.col).width || 10;
      const fontSize = cell.font?.size || st.FONT_SIZE;
      const lineCount = text.split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / Math.max(1, width * 1.2))), 0);
      rowHeight = Math.max(rowHeight, lineCount * Math.max(18, fontSize * st.ROW_HEIGHT_FONT_MULTIPLIER));
    });
    row.height = Math.min(st.MAX_ROW_HEIGHT, rowHeight);
  });
}

function formatTableRangeImpl(sheet: ExcelJS.Worksheet, startRow: number, endRow: number, startCol: number, endCol: number, headerRow: number, st: ExcelStyleConfig): void {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let col = startCol; col <= endCol; col += 1) {
      styleCellImpl(sheet.getCell(row, col), {
        border: true,
        bold: row === headerRow,
        fill: row === headerRow ? st.HEADER_FILL : undefined,
      }, st);
    }
  }
}

export function productFormula(addresses: string[], exponent: number): string {
  return `(${addresses.join("*")})^(1/${exponent})`;
}

export function absoluteRef(address: string): string {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  return match ? `$${match[1]}$${match[2]}` : address;
}

export function weightedSumFormula(valueAddresses: string[], weightAddresses: string[]): string {
  return valueAddresses.map((address, index) => `${address}*${weightAddresses[index]}`).join("+");
}

export function listFormulaRange(sheetName: string, addresses: string[]): string {
  return addresses.map((address) => `'${sheetName}'!${address}`).join(",");
}

function addAhpMatrixBlockImpl(sheet: ExcelJS.Worksheet, startRow: number, title: string, labels: string[], matrix: number[][], randomIndex: number, st: ExcelStyleConfig): AhpBlockResult {
  const labelStyle: CellStyleOptions = { bold: true, border: true, fill: st.HEADER_FILL, align: "left" };
  setValueImpl(sheet, `A${startRow}`, title, { ...labelStyle, size: st.HEADER_TITLE_SIZE }, st);
  sheet.getRow(startRow).height = Math.max(sheet.getRow(startRow).height || st.DEFAULT_ROW_HEIGHT, 34);
  labels.forEach((label, idx) => {
    setValueImpl(sheet, `${String.fromCharCode(66 + idx)}${startRow + 2}`, label, { bold: true, border: true, fill: st.HEADER_FILL }, st);
    setValueImpl(sheet, `A${startRow + 3 + idx}`, label, { bold: true, border: true, fill: st.HEADER_FILL }, st);
  });
  setValueImpl(sheet, `${String.fromCharCode(66 + labels.length)}${startRow + 2}`, "G_i", { bold: true, border: true, fill: st.HEADER_FILL }, st);
  setValueImpl(sheet, `${String.fromCharCode(67 + labels.length)}${startRow + 2}`, "w_i", { bold: true, border: true, fill: st.HEADER_FILL }, st);

  for (let r = 0; r < labels.length; r += 1) {
    const row = startRow + 3 + r;
    const cells: string[] = [];
    for (let c = 0; c < labels.length; c += 1) {
      const address = `${String.fromCharCode(66 + c)}${row}`;
      cells.push(address);
      setValueImpl(sheet, address, matrix[r][c], { border: true, numFmt: "0.0000" }, st);
    }
    const gCol = String.fromCharCode(66 + labels.length);
    const wCol = String.fromCharCode(67 + labels.length);
    setFormulaImpl(sheet, `${gCol}${row}`, productFormula(cells, labels.length), { border: true, numFmt: "0.0000" }, st);
    setFormulaImpl(sheet, `${wCol}${row}`, `${gCol}${row}/SUM(${gCol}${startRow + 3}:${gCol}${startRow + 2 + labels.length})`, { border: true, numFmt: "0.0000" }, st);
  }

  const statsRow = startRow + 4 + labels.length;
  setValueImpl(sheet, `A${statsRow}`, "Сумма весов:", labelStyle, st);
  const weightsCol = String.fromCharCode(67 + labels.length);
  setFormulaImpl(sheet, `B${statsRow}`, `SUM(${weightsCol}${startRow + 3}:${weightsCol}${startRow + 2 + labels.length})`, { border: true, numFmt: "0.0000" }, st);

  const colSumsRow = statsRow + 2;
  setValueImpl(sheet, `A${colSumsRow}`, "Суммы столбцов:", labelStyle, st);
  for (let c = 0; c < labels.length; c += 1) {
    const letter = String.fromCharCode(66 + c);
    setFormulaImpl(sheet, `${letter}${colSumsRow}`, `SUM(${letter}${startRow + 3}:${letter}${startRow + 2 + labels.length})`, { border: true, numFmt: "0.0000" }, st);
  }
  const colSumAddresses = labels.map((_, index) => `${String.fromCharCode(66 + index)}${colSumsRow}`);
  const weightAddresses = labels.map((_, index) => `${weightsCol}${startRow + 3 + index}`);
  setValueImpl(sheet, `A${colSumsRow + 1}`, "λ_max =", labelStyle, st);
  setFormulaImpl(sheet, `B${colSumsRow + 1}`, weightedSumFormula(colSumAddresses, weightAddresses), { border: true, numFmt: "0.0000" }, st);
  setValueImpl(sheet, `A${colSumsRow + 2}`, "ИС =", labelStyle, st);
  setFormulaImpl(sheet, `B${colSumsRow + 2}`, `(B${colSumsRow + 1}-${labels.length})/${labels.length - 1}`, { border: true, numFmt: "0.0000" }, st);
  setValueImpl(sheet, `A${colSumsRow + 3}`, `ИС_сл (n=${labels.length}) =`, labelStyle, st);
  setValueImpl(sheet, `B${colSumsRow + 3}`, randomIndex, { border: true, numFmt: "0.00" }, st);
  setValueImpl(sheet, `A${colSumsRow + 4}`, "ОС =", labelStyle, st);
  setFormulaImpl(sheet, `B${colSumsRow + 4}`, `B${colSumsRow + 2}/B${colSumsRow + 3}`, { border: true, numFmt: "0.0000" }, st);
  setValueImpl(sheet, `A${colSumsRow + 5}`, "ОС% =", labelStyle, st);
  setFormulaImpl(sheet, `B${colSumsRow + 5}`, `B${colSumsRow + 4}*100`, { border: true, numFmt: "0.00" }, st);

  return { weightColumn: weightsCol, weightStartRow: startRow + 3, weightEndRow: startRow + 2 + labels.length, osCell: `B${colSumsRow + 4}`, nextRow: colSumsRow + 7 };
}

function addContentsSheetImpl(workbook: ExcelJS.Workbook, sheetNames: string[], st: ExcelStyleConfig): ExcelJS.Worksheet {
  const sheet = workbook.addWorksheet("Содержание", { properties: { tabColor: { argb: st.TAB_COLOR } } });
  applyBaseSheetSetupImpl(sheet, [40, 70], st);
  mergeTitleImpl(sheet, "A1:B1", "СОДЕРЖАНИЕ", st);
  setValueImpl(sheet, "A3", "Лист", { bold: true, border: true, fill: st.HEADER_FILL, align: "center" }, st);
  setValueImpl(sheet, "B3", "Переход", { bold: true, border: true, fill: st.HEADER_FILL, align: "center" }, st);
  sheetNames.forEach((name, index) => {
    const row = 4 + index;
    setValueImpl(sheet, `A${row}`, name, { border: true, align: "left" }, st);
    const link = sheet.getCell(`B${row}`);
    link.value = { text: `Перейти к листу "${name}"`, hyperlink: `#'${name}'!A1` };
    styleCellImpl(link, { border: true, align: "left", color: "0563C1" }, st);
  });
  setFreezeImpl(sheet, 3, "A4");
  return sheet;
}

export function createExcelGost(opts: Partial<ExcelStyleConfig> = {}): ExcelGostInstance {
  const st = createExcelStyle(opts);
  const HEADER_STYLE: CellStyleOptions = { bold: true, border: true, fill: st.HEADER_FILL, align: "center" };
  const RESULT_STYLE: CellStyleOptions = { bold: true, border: "medium", fill: st.RESULT_FILL, align: "center" };
  return {
    style: st,
    HEADER_STYLE,
    RESULT_STYLE,
    styleCell: (cell, options = {}) => styleCellImpl(cell, options, st),
    setValue: (sheet, address, value, options = {}) => setValueImpl(sheet, address, value, options, st),
    setFormula: (sheet, address, formula, options = {}) => setFormulaImpl(sheet, address, formula, options, st),
    applyBaseSheetSetup: (sheet, widths) => applyBaseSheetSetupImpl(sheet, widths, st),
    mergeTitle: (sheet, range, text) => mergeTitleImpl(sheet, range, text, st),
    setFreeze: (sheet, ySplit, activeCell) => setFreezeImpl(sheet, ySplit, activeCell),
    applyZebraStriping: (sheet, sr, er, sc, ec, hr) => applyZebraStripingImpl(sheet, sr, er, sc, ec, hr, st),
    getCellText: (cell) => getCellTextImpl(cell),
    autoFitSheet: (sheet) => autoFitSheetImpl(sheet, st),
    formatTableRange: (sheet, sr, er, sc, ec, hr) => formatTableRangeImpl(sheet, sr, er, sc, ec, hr, st),
    addAhpMatrixBlock: (sheet, sr, title, labels, matrix, ri) => addAhpMatrixBlockImpl(sheet, sr, title, labels, matrix, ri, st),
    addContentsSheet: (wb, names) => addContentsSheetImpl(wb, names, st),
    productFormula,
    absoluteRef,
    weightedSumFormula,
    listFormulaRange,
  };
}

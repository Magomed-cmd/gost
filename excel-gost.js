"use strict";

function createExcelStyle(opts = {}) {
    return {
        HEADER_FILL: opts.headerFill ?? "FFF2CC",
        RESULT_FILL: opts.resultFill ?? "D9EAD3",
        ALT_ROW_FILL: opts.altRowFill ?? "FFF9F0",
        REGRET_FILL: opts.regretFill ?? "FCE4D6",
        FONT_NAME: opts.fontName ?? "Times New Roman",
        FONT_SIZE: opts.fontSize ?? 14,
        TAB_COLOR: opts.tabColor ?? "FFF2CC",
        DEFAULT_ROW_HEIGHT: opts.defaultRowHeight ?? 20,
        PAGE_ORIENTATION: opts.pageOrientation ?? "landscape",
        FIT_TO_PAGE: opts.fitToPage ?? true,
        FIT_TO_WIDTH: opts.fitToWidth ?? 1,
        FIT_TO_HEIGHT: opts.fitToHeight ?? 0,
        MAX_COLUMN_WIDTH: opts.maxColumnWidth ?? 80,
        MAX_ROW_HEIGHT: opts.maxRowHeight ?? 120,
        ROW_HEIGHT_FONT_MULTIPLIER: opts.rowHeightFontMultiplier ?? 1.9,
        HEADER_TITLE_SIZE: opts.headerTitleSize ?? 14,
        BORDER_STYLE: opts.borderStyle ?? "thin",
    };
}

function styleCell(cell, options = {}, style = {}) {
    cell.font = {
        name: options.fontName ?? style.FONT_NAME ?? "Times New Roman",
        size: options.size ?? style.FONT_SIZE ?? 12,
        bold: Boolean(options.bold),
        color: options.color ? { argb: options.color } : undefined,
    };
    cell.alignment = {
        vertical: "middle",
        horizontal: options.align || "center",
        wrapText: options.wrapText !== false,
    };
    if (options.fill) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: options.fill } };
    }
    if (options.border) {
        const borderStyle = typeof options.border === "string" ? options.border : (style.BORDER_STYLE ?? "thin");
        cell.border = {
            top: { style: borderStyle },
            left: { style: borderStyle },
            bottom: { style: borderStyle },
            right: { style: borderStyle },
        };
    }
    if (options.numFmt) {
        cell.numFmt = options.numFmt;
    }
}

function setValue(sheet, address, value, options = {}, style = {}) {
    const cell = sheet.getCell(address);
    cell.value = value;
    styleCell(cell, options, style);
    return cell;
}

function setFormula(sheet, address, formula, options = {}, style = {}) {
    const cell = sheet.getCell(address);
    cell.value = { formula };
    styleCell(cell, options, style);
    return cell;
}

function applyBaseSheetSetup(sheet, widths, style = {}) {
    sheet.properties.defaultRowHeight = style.DEFAULT_ROW_HEIGHT ?? 20;
    sheet.pageSetup.orientation = style.PAGE_ORIENTATION ?? "landscape";
    sheet.pageSetup.fitToPage = style.FIT_TO_PAGE ?? true;
    sheet.pageSetup.fitToWidth = style.FIT_TO_WIDTH ?? 1;
    sheet.pageSetup.fitToHeight = style.FIT_TO_HEIGHT ?? 0;
    sheet._baseWidths = widths;
    widths.forEach((width, index) => {
        sheet.getColumn(index + 1).width = width;
    });
}

function mergeTitle(sheet, range, text, style = {}) {
    sheet.mergeCells(range);
    const cell = sheet.getCell(range.split(":")[0]);
    cell.value = text;
    styleCell(cell, { bold: true, size: style.HEADER_TITLE_SIZE ?? 14, align: "center" }, style);
}

function setFreeze(sheet, ySplit, activeCell = "A2") {
    void sheet;
    void ySplit;
    void activeCell;
}

function applyZebraStriping(sheet, startRow, endRow, startCol, endCol, headerRows = new Set(), style = {}) {
    for (let row = startRow; row <= endRow; row += 1) {
        if (headerRows.has(row) || row % 2 !== 0) continue;
        for (let col = startCol; col <= endCol; col += 1) {
            const cell = sheet.getCell(row, col);
            const existingFill = cell.fill?.fgColor?.argb;
            if (existingFill === style.HEADER_FILL || existingFill === style.RESULT_FILL) continue;
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: style.ALT_ROW_FILL ?? "FFF9F0" } };
        }
    }
}

function getCellText(cell) {
    if (!cell || cell.value === null || cell.value === undefined) return "";
    if (typeof cell.value === "object" && cell.value.formula) return "";
    if (typeof cell.value === "object" && cell.value.text) return String(cell.value.text);
    return String(cell.text || cell.value);
}

function autoFitSheet(sheet, style = {}) {
    const baseWidths = sheet._baseWidths || [];
    const maxColumn = Math.max(sheet.columnCount, baseWidths.length);
    const computedWidths = Array.from({ length: maxColumn }, (_, index) => baseWidths[index] || 10);

    for (let columnIndex = 1; columnIndex <= maxColumn; columnIndex += 1) {
        sheet.eachRow({ includeEmpty: false }, (row) => {
            const text = getCellText(row.getCell(columnIndex)).trim();
            if (!text) return;
            const longestLine = text.split(/\r?\n/).reduce((max, line) => Math.max(max, line.length), 0);
            computedWidths[columnIndex - 1] = Math.min(style.MAX_COLUMN_WIDTH ?? 80, Math.max(computedWidths[columnIndex - 1], longestLine + 2));
        });
    }
    computedWidths.forEach((width, index) => {
        sheet.getColumn(index + 1).width = width;
    });

    sheet.eachRow({ includeEmpty: false }, (row) => {
        let rowHeight = sheet.properties.defaultRowHeight || 20;
        row.eachCell({ includeEmpty: false }, (cell) => {
            const text = getCellText(cell).trim();
            if (!text) return;
            const width = sheet.getColumn(cell.col).width || 10;
            const fontSize = cell.font?.size || 12;
            const lineCount = text.split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / Math.max(1, width * 1.2))), 0);
            rowHeight = Math.max(rowHeight, lineCount * Math.max(18, fontSize * (style.ROW_HEIGHT_FONT_MULTIPLIER ?? 1.9)));
        });
        row.height = Math.min(style.MAX_ROW_HEIGHT ?? 120, rowHeight);
    });
}

function formatTableRange(sheet, startRow, endRow, startCol, endCol, headerRow, style = {}) {
    for (let row = startRow; row <= endRow; row += 1) {
        for (let col = startCol; col <= endCol; col += 1) {
            styleCell(sheet.getCell(row, col), {
                border: true,
                bold: row === headerRow,
                fill: row === headerRow ? style.HEADER_FILL : undefined,
            }, style);
        }
    }
}

function productFormula(addresses, exponent) {
    return `(${addresses.join("*")})^(1/${exponent})`;
}

function absoluteRef(address) {
    const match = address.match(/^([A-Z]+)(\d+)$/);
    if (!match) return address;
    return `$${match[1]}$${match[2]}`;
}

function weightedSumFormula(valueAddresses, weightAddresses) {
    return valueAddresses.map((address, index) => `${address}*${weightAddresses[index]}`).join("+");
}

function listFormulaRange(sheetName, addresses) {
    return addresses.map((address) => `'${sheetName}'!${address}`).join(",");
}

function addAhpMatrixBlock(sheet, startRow, title, labels, matrix, randomIndex, style = {}) {
  const labelStyle = { bold: true, border: true, fill: style.HEADER_FILL, align: "left" };
  setValue(sheet, `A${startRow}`, title, { ...labelStyle, size: 14 }, style);
    sheet.getRow(startRow).height = Math.max(sheet.getRow(startRow).height || 20, 34);
    labels.forEach((label, idx) => {
        setValue(sheet, `${String.fromCharCode(66 + idx)}${startRow + 2}`, label, { bold: true, border: true, fill: style.HEADER_FILL }, style);
        setValue(sheet, `A${startRow + 3 + idx}`, label, { bold: true, border: true, fill: style.HEADER_FILL }, style);
    });
    setValue(sheet, `${String.fromCharCode(66 + labels.length)}${startRow + 2}`, "G_i", { bold: true, border: true, fill: style.HEADER_FILL }, style);
    setValue(sheet, `${String.fromCharCode(67 + labels.length)}${startRow + 2}`, "w_i", { bold: true, border: true, fill: style.HEADER_FILL }, style);

    for (let r = 0; r < labels.length; r += 1) {
        const row = startRow + 3 + r;
        const valueCells = [];
        for (let c = 0; c < labels.length; c += 1) {
            const colLetter = String.fromCharCode(66 + c);
            const address = `${colLetter}${row}`;
            valueCells.push(address);
            setValue(sheet, address, matrix[r][c], { border: true, numFmt: "0.0000" }, style);
        }
        const gCol = String.fromCharCode(66 + labels.length);
        const wCol = String.fromCharCode(67 + labels.length);
        setFormula(sheet, `${gCol}${row}`, productFormula(valueCells, labels.length), { border: true, numFmt: "0.0000" }, style);
        setFormula(sheet, `${wCol}${row}`, `${gCol}${row}/SUM(${gCol}${startRow + 3}:${gCol}${startRow + 2 + labels.length})`, { border: true, numFmt: "0.0000" }, style);
    }

    const statsRow = startRow + 4 + labels.length;
  setValue(sheet, `A${statsRow}`, "Сумма весов:", labelStyle, style);
    setFormula(sheet, `B${statsRow}`, `SUM(${String.fromCharCode(67 + labels.length)}${startRow + 3}:${String.fromCharCode(67 + labels.length)}${startRow + 2 + labels.length})`, { border: true, numFmt: "0.0000" }, style);

    const colSumsRow = statsRow + 2;
  setValue(sheet, `A${colSumsRow}`, "Суммы столбцов:", labelStyle, style);
    for (let c = 0; c < labels.length; c += 1) {
        const colLetter = String.fromCharCode(66 + c);
        setFormula(sheet, `${colLetter}${colSumsRow}`, `SUM(${colLetter}${startRow + 3}:${colLetter}${startRow + 2 + labels.length})`, { border: true, numFmt: "0.0000" }, style);
    }
  setValue(sheet, `A${colSumsRow + 1}`, "λ_max =", labelStyle, style);
    const weightsCol = String.fromCharCode(67 + labels.length);
    const colSumAddresses = labels.map((_, index) => `${String.fromCharCode(66 + index)}${colSumsRow}`);
    const weightAddresses = labels.map((_, index) => `${weightsCol}${startRow + 3 + index}`);
    setFormula(sheet, `B${colSumsRow + 1}`, weightedSumFormula(colSumAddresses, weightAddresses), { border: true, numFmt: "0.0000" }, style);
  setValue(sheet, `A${colSumsRow + 2}`, "ИС =", labelStyle, style);
    setFormula(sheet, `B${colSumsRow + 2}`, `(B${colSumsRow + 1}-${labels.length})/${labels.length - 1}`, { border: true, numFmt: "0.0000" }, style);
  setValue(sheet, `A${colSumsRow + 3}`, `ИС_сл (n=${labels.length}) =`, labelStyle, style);
    setValue(sheet, `B${colSumsRow + 3}`, randomIndex, { border: true, numFmt: "0.00" }, style);
  setValue(sheet, `A${colSumsRow + 4}`, "ОС =", labelStyle, style);
    setFormula(sheet, `B${colSumsRow + 4}`, `B${colSumsRow + 2}/B${colSumsRow + 3}`, { border: true, numFmt: "0.0000" }, style);
  setValue(sheet, `A${colSumsRow + 5}`, "ОС% =", labelStyle, style);
    setFormula(sheet, `B${colSumsRow + 5}`, `B${colSumsRow + 4}*100`, { border: true, numFmt: "0.00" }, style);

    return {
        weightColumn: weightsCol,
        weightStartRow: startRow + 3,
        weightEndRow: startRow + 2 + labels.length,
        osCell: `B${colSumsRow + 4}`,
        nextRow: colSumsRow + 7,
    };
}

function addContentsSheet(workbook, sheetNames, style = {}) {
    const sheet = workbook.addWorksheet("Содержание", { properties: { tabColor: { argb: style.TAB_COLOR ?? "FFF2CC" } } });
    applyBaseSheetSetup(sheet, [40, 70], style);
    mergeTitle(sheet, "A1:B1", "СОДЕРЖАНИЕ", style);
    setValue(sheet, "A3", "Лист", { bold: true, border: true, fill: style.HEADER_FILL, align: "center" }, style);
    setValue(sheet, "B3", "Переход", { bold: true, border: true, fill: style.HEADER_FILL, align: "center" }, style);
    sheetNames.forEach((name, index) => {
        const row = 4 + index;
        setValue(sheet, `A${row}`, name, { border: true, align: "left" }, style);
        const linkCell = sheet.getCell(`B${row}`);
        linkCell.value = { text: `Перейти к листу "${name}"`, hyperlink: `#'${name}'!A1` };
        styleCell(linkCell, { border: true, align: "left", color: "0563C1" }, style);
    });
    setFreeze(sheet, 3, "A4");
    return sheet;
}

module.exports = {
    createExcelStyle,
    styleCell, setValue, setFormula,
    applyBaseSheetSetup, mergeTitle, setFreeze,
    applyZebraStriping, getCellText, autoFitSheet,
    formatTableRange, addAhpMatrixBlock, addContentsSheet,
    productFormula, absoluteRef, weightedSumFormula, listFormulaRange,
};

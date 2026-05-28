"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.TableOfContents = exports.DEFAULT_DOCX_STYLE = void 0;
exports.createDocxStyle = createDocxStyle;
exports.toMathComponents = toMathComponents;
exports.mathExpr = mathExpr;
exports.mathFraction = mathFraction;
exports.mathSub = mathSub;
exports.mathSup = mathSup;
exports.mathSubSup = mathSubSup;
exports.mathParen = mathParen;
exports.mathBracket = mathBracket;
exports.mathBrace = mathBrace;
exports.mathRoot = mathRoot;
exports.mathSum = mathSum;
exports.mathIntegral = mathIntegral;
exports.createDocxGost = createDocxGost;
const fs_1 = __importDefault(require("fs"));
const docx_1 = require("docx");
Object.defineProperty(exports, "TableOfContents", { enumerable: true, get: function () { return docx_1.TableOfContents; } });
exports.DEFAULT_DOCX_STYLE = {
    FONT: "Times New Roman",
    SIZE: 28,
    SIZE_TITLE: 32,
    SIZE_HEADER: 24,
    SIZE_CODE: 24,
    INDENT: 709,
    LINE: 360,
    CONTENT_WIDTH: 9355,
    CODE_FONT: "Courier New",
    TABLE_CELL_LINE: 276,
    TABLE_CELL_MARGINS: { top: 30, bottom: 30, left: 60, right: 60 },
    HEADING_COLOR: "000000",
    CAPTION_BEFORE: 60,
    CAPTION_AFTER: 60,
    TABLE_CAPTION_BEFORE: 120,
    TABLE_CAPTION_AFTER: 60,
    IMAGE_BEFORE: 120,
    IMAGE_AFTER: 60,
    FORMULA_BEFORE: 80,
    FORMULA_AFTER: 80,
    PAGE: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1134, bottom: 1134, left: 1701, right: 850 },
    },
    BORDER_SINGLE: { style: docx_1.BorderStyle.SINGLE, size: 1, color: "000000" },
    BORDERS_ALL: {
        top: { style: docx_1.BorderStyle.SINGLE, size: 1, color: "000000" },
        bottom: { style: docx_1.BorderStyle.SINGLE, size: 1, color: "000000" },
        left: { style: docx_1.BorderStyle.SINGLE, size: 1, color: "000000" },
        right: { style: docx_1.BorderStyle.SINGLE, size: 1, color: "000000" },
    },
    TITLE_TAB_STOPS: [
        { type: docx_1.TabStopType.LEFT, position: 4500 },
        { type: docx_1.TabStopType.RIGHT, position: 9355 },
    ],
};
function createDocxStyle(opts = {}) {
    const borderSingle = opts.BORDER_SINGLE ?? exports.DEFAULT_DOCX_STYLE.BORDER_SINGLE;
    return {
        ...exports.DEFAULT_DOCX_STYLE,
        ...opts,
        FONT: opts.font ?? opts.FONT ?? exports.DEFAULT_DOCX_STYLE.FONT,
        SIZE: opts.size ?? opts.SIZE ?? exports.DEFAULT_DOCX_STYLE.SIZE,
        SIZE_TITLE: opts.sizeTitle ?? opts.SIZE_TITLE ?? exports.DEFAULT_DOCX_STYLE.SIZE_TITLE,
        SIZE_HEADER: opts.sizeHeader ?? opts.SIZE_HEADER ?? exports.DEFAULT_DOCX_STYLE.SIZE_HEADER,
        SIZE_CODE: opts.sizeCode ?? opts.SIZE_CODE ?? exports.DEFAULT_DOCX_STYLE.SIZE_CODE,
        INDENT: opts.indent ?? opts.INDENT ?? exports.DEFAULT_DOCX_STYLE.INDENT,
        LINE: opts.line ?? opts.LINE ?? exports.DEFAULT_DOCX_STYLE.LINE,
        CONTENT_WIDTH: opts.contentWidth ?? opts.CONTENT_WIDTH ?? exports.DEFAULT_DOCX_STYLE.CONTENT_WIDTH,
        CODE_FONT: opts.codeFont ?? opts.CODE_FONT ?? exports.DEFAULT_DOCX_STYLE.CODE_FONT,
        PAGE: {
            size: { ...exports.DEFAULT_DOCX_STYLE.PAGE.size, ...(opts.page?.size ?? opts.PAGE?.size ?? {}) },
            margin: { ...exports.DEFAULT_DOCX_STYLE.PAGE.margin, ...(opts.page?.margin ?? opts.PAGE?.margin ?? {}) },
        },
        BORDER_SINGLE: borderSingle,
        BORDERS_ALL: opts.BORDERS_ALL ?? {
            top: borderSingle,
            bottom: borderSingle,
            left: borderSingle,
            right: borderSingle,
        },
        TITLE_TAB_STOPS: opts.TITLE_TAB_STOPS ?? exports.DEFAULT_DOCX_STYLE.TITLE_TAB_STOPS,
    };
}
function isMathLike(value) {
    return value instanceof docx_1.Math
        || (typeof value === "object" && value !== null && value.constructor.name.startsWith("Math"));
}
function runImpl(st, text, opts = {}) {
    return new docx_1.TextRun({
        text: String(text ?? ""),
        font: opts.font ?? st.FONT,
        size: opts.size ?? st.SIZE,
        bold: !!opts.bold,
        italics: !!opts.italics,
        allCaps: !!opts.allCaps,
        underline: opts.underline ? (opts.underline === true ? {} : opts.underline) : undefined,
        color: opts.color,
    });
}
function paragraphOptions(st, opts = {}, children = []) {
    return {
        heading: undefined,
        alignment: opts.align || docx_1.AlignmentType.JUSTIFIED,
        spacing: { line: opts.line ?? st.LINE, before: opts.before ?? 0, after: opts.after ?? 0 },
        indent: opts.noIndent ? undefined : { firstLine: opts.indent ?? st.INDENT },
        keepNext: !!opts.keepNext,
        keepLines: opts.keepLines !== undefined ? opts.keepLines : true,
        children,
    };
}
function paragraphChildren(st, content, opts = {}) {
    if (typeof content === "string" || typeof content === "number") {
        return [runImpl(st, content, opts)];
    }
    return content.map((item) => {
        if (typeof item === "string" || typeof item === "number") {
            return runImpl(st, item, opts);
        }
        return runImpl(st, item.text, {
            ...opts,
            size: item.size ?? opts.size,
            bold: item.bold ?? opts.bold,
            italics: item.italics ?? opts.italics,
            allCaps: item.allCaps ?? opts.allCaps,
            underline: item.underline ?? opts.underline,
            color: item.color ?? opts.color,
            font: item.font ?? opts.font,
        });
    });
}
function paragraphImpl(st, content, opts = {}) {
    return new docx_1.Paragraph(paragraphOptions(st, opts, paragraphChildren(st, content, opts)));
}
function centeredImpl(st, content, opts = {}) {
    return paragraphImpl(st, content, { ...opts, align: docx_1.AlignmentType.CENTER, noIndent: true });
}
function titleLine(st, content, opts = {}) {
    return centeredImpl(st, content, {
        ...opts,
        size: opts.size ?? st.SIZE,
        line: opts.line ?? st.LINE,
        before: opts.before ?? 0,
        after: opts.after ?? 0,
    });
}
function defaultStyles(st) {
    return {
        default: { document: { run: { font: st.FONT, size: st.SIZE } } },
        paragraphStyles: [
            {
                id: "Heading1",
                name: "Heading 1",
                basedOn: "Normal",
                next: "Normal",
                run: { font: st.FONT, size: st.SIZE, bold: true, color: st.HEADING_COLOR },
                paragraph: { alignment: docx_1.AlignmentType.CENTER, spacing: { before: 240, after: 120 } },
            },
            {
                id: "Heading2",
                name: "Heading 2",
                basedOn: "Normal",
                next: "Normal",
                run: { font: st.FONT, size: st.SIZE, bold: true, color: st.HEADING_COLOR },
                paragraph: { spacing: { before: 200, after: 80 } },
            },
            {
                id: "Heading3",
                name: "Heading 3",
                basedOn: "Normal",
                next: "Normal",
                run: { font: st.FONT, size: st.SIZE, bold: true, italics: true, color: st.HEADING_COLOR },
                paragraph: { spacing: { before: 140, after: 60 } },
            },
        ],
    };
}
function h1Impl(st, text, opts = {}) {
    return new docx_1.Paragraph({
        ...paragraphOptions(st, { ...opts, align: docx_1.AlignmentType.CENTER, noIndent: true, before: 240, after: 120, keepNext: true }, [
            runImpl(st, text.toUpperCase(), { ...opts, bold: true, color: st.HEADING_COLOR }),
        ]),
        heading: docx_1.HeadingLevel.HEADING_1,
    });
}
function h2Impl(st, text, opts = {}) {
    return new docx_1.Paragraph({
        ...paragraphOptions(st, { ...opts, before: 200, after: 80, keepNext: true }, [
            runImpl(st, text, { ...opts, bold: true, color: st.HEADING_COLOR }),
        ]),
        heading: docx_1.HeadingLevel.HEADING_2,
    });
}
function h3Impl(st, text, opts = {}) {
    return new docx_1.Paragraph({
        ...paragraphOptions(st, { ...opts, before: 140, after: 60, keepNext: true }, [
            runImpl(st, text, { ...opts, bold: true, italics: true, color: st.HEADING_COLOR }),
        ]),
        heading: docx_1.HeadingLevel.HEADING_3,
    });
}
function blankImpl(st, opts = {}) {
    return new docx_1.Paragraph({ spacing: { line: opts.line ?? st.LINE }, children: [runImpl(st, "", opts)] });
}
function pageBreakImpl() {
    return new docx_1.Paragraph({ children: [new docx_1.PageBreak()] });
}
function captionImpl(st, text, opts = {}) {
    return centeredImpl(st, text, {
        ...opts,
        before: opts.before ?? st.CAPTION_BEFORE,
        after: opts.after ?? st.CAPTION_AFTER,
        size: opts.size ?? st.SIZE,
    });
}
function tableCaptionImpl(st, text, opts = {}) {
    return paragraphImpl(st, text, {
        ...opts,
        align: docx_1.AlignmentType.LEFT,
        noIndent: true,
        before: opts.before ?? st.TABLE_CAPTION_BEFORE,
        after: opts.after ?? st.TABLE_CAPTION_AFTER,
        size: opts.size ?? st.SIZE,
    });
}
function createCaptionCounterImpl(st, prefix, opts = {}) {
    let n = opts.start ?? 1;
    const fn = opts.table ? tableCaptionImpl : captionImpl;
    return {
        peek: () => n,
        caption: (text) => fn(st, `${prefix} ${n++} – ${text}`, opts),
    };
}
function placeholderImpl(st, label, opts = {}) {
    const border = { style: docx_1.BorderStyle.SINGLE, size: 4, color: "AAAAAA" };
    return new docx_1.Paragraph({
        ...paragraphOptions(st, {
            ...opts,
            align: docx_1.AlignmentType.CENTER,
            noIndent: true,
            before: opts.before ?? 120,
            after: opts.after ?? 60,
        }, [runImpl(st, label, { ...opts, italics: true, color: opts.color ?? "888888" })]),
        border: { top: border, bottom: border, left: border, right: border },
    });
}
function imageBlockImpl(st, imagePath, width, height, opts = {}) {
    const data = Buffer.isBuffer(imagePath) ? imagePath : fs_1.default.readFileSync(imagePath);
    return new docx_1.Paragraph({
        alignment: opts.align || docx_1.AlignmentType.CENTER,
        spacing: {
            line: opts.line ?? st.LINE,
            before: opts.before ?? st.IMAGE_BEFORE,
            after: opts.after ?? st.IMAGE_AFTER,
        },
        keepNext: !!opts.keepNext,
        keepLines: opts.keepLines !== undefined ? opts.keepLines : true,
        children: [new docx_1.ImageRun({ data, type: opts.imageType ?? "png", transformation: { width, height } })],
    });
}
function codeLineImpl(st, text, opts = {}) {
    return paragraphImpl(st, [{ text, font: opts.font ?? st.CODE_FONT, size: opts.size ?? st.SIZE_CODE }], {
        ...opts,
        noIndent: true,
        line: opts.line ?? st.LINE,
    });
}
function toMathComponents(value) {
    if (value === undefined || value === null)
        return [];
    if (Array.isArray(value))
        return value.flatMap((item) => toMathComponents(item));
    if (typeof value === "string" || typeof value === "number")
        return [new docx_1.MathRun(String(value))];
    return [value];
}
function mathExpr(...parts) {
    return new docx_1.Math({ children: parts.flatMap((part) => toMathComponents(part)) });
}
function mathFraction(numerator, denominator) {
    return new docx_1.MathFraction({ numerator: toMathComponents(numerator), denominator: toMathComponents(denominator) });
}
function mathSub(base, subScript) {
    return new docx_1.MathSubScript({ children: toMathComponents(base), subScript: toMathComponents(subScript) });
}
function mathSup(base, superScript) {
    return new docx_1.MathSuperScript({ children: toMathComponents(base), superScript: toMathComponents(superScript) });
}
function mathSubSup(base, subScript, superScript) {
    return new docx_1.MathSubSuperScript({
        children: toMathComponents(base),
        subScript: toMathComponents(subScript),
        superScript: toMathComponents(superScript),
    });
}
function mathParen(...parts) {
    return new docx_1.MathRoundBrackets({ children: parts.flatMap((part) => toMathComponents(part)) });
}
function mathBracket(...parts) {
    return new docx_1.MathSquareBrackets({ children: parts.flatMap((part) => toMathComponents(part)) });
}
function mathBrace(...parts) {
    return new docx_1.MathCurlyBrackets({ children: parts.flatMap((part) => toMathComponents(part)) });
}
function mathRoot(content, degree) {
    return new docx_1.MathRadical({ children: toMathComponents(content), degree: degree === undefined ? undefined : toMathComponents(degree) });
}
function mathSum(content, opts = {}) {
    return new docx_1.MathSum({
        children: toMathComponents(content),
        subScript: opts.subScript ? toMathComponents(opts.subScript) : undefined,
        superScript: opts.superScript ? toMathComponents(opts.superScript) : undefined,
    });
}
function mathIntegral(content, opts = {}) {
    return new docx_1.MathIntegral({
        children: toMathComponents(content),
        subScript: opts.subScript ? toMathComponents(opts.subScript) : undefined,
        superScript: opts.superScript ? toMathComponents(opts.superScript) : undefined,
    });
}
function formulaMathImpl(st, content, opts = {}) {
    return new docx_1.Paragraph(paragraphOptions(st, {
        ...opts,
        align: docx_1.AlignmentType.CENTER,
        noIndent: true,
        before: opts.before ?? st.FORMULA_BEFORE,
        after: opts.after ?? st.FORMULA_AFTER,
    }, [mathExpr(content)]));
}
function formulaInlineImpl(st, label, mathContent, opts = {}) {
    return new docx_1.Paragraph(paragraphOptions(st, opts, [runImpl(st, `${label} `, opts), mathExpr(mathContent)]));
}
function paragraphWithMathImpl(st, parts, opts = {}) {
    const children = parts.map((part) => {
        if (typeof part === "string" || typeof part === "number")
            return runImpl(st, part, opts);
        if (part instanceof docx_1.Math)
            return part;
        if (Array.isArray(part) || isMathLike(part))
            return mathExpr(part);
        const desc = part;
        return runImpl(st, desc.text, {
            ...opts,
            size: desc.size ?? opts.size,
            bold: desc.bold ?? opts.bold,
            italics: desc.italics ?? opts.italics,
            allCaps: desc.allCaps ?? opts.allCaps,
            underline: desc.underline ?? opts.underline,
            color: desc.color ?? opts.color,
            font: desc.font ?? opts.font,
        });
    });
    return new docx_1.Paragraph(paragraphOptions(st, opts, children));
}
function normalizeCell(cell) {
    if (typeof cell === "object" && cell !== null && !Array.isArray(cell) && !isMathLike(cell) && "text" in cell) {
        return cell;
    }
    return { text: cell };
}
function tableColumnWidths(st, count, widths) {
    if (Array.isArray(widths) && widths.length > 0)
        return widths;
    const base = Math.floor(st.CONTENT_WIDTH / count);
    const result = Array(count).fill(base);
    result[result.length - 1] += st.CONTENT_WIDTH - result.reduce((sum, width) => sum + width, 0);
    return result;
}
function tableCellImpl(st, cell, width, opts) {
    const normalized = normalizeCell(cell);
    const content = normalized.text ?? "";
    const align = normalized.align ?? opts.align;
    const bold = normalized.bold !== undefined ? normalized.bold : opts.bold;
    const children = Array.isArray(content) || isMathLike(content)
        ? [paragraphWithMathImpl(st, Array.isArray(content) ? content : [content], {
                align,
                noIndent: true,
                line: opts.cellLine ?? st.TABLE_CELL_LINE,
                size: opts.fontSize,
            })]
        : [new docx_1.Paragraph({
                alignment: align,
                spacing: { line: opts.cellLine ?? st.TABLE_CELL_LINE, before: 0, after: 0 },
                children: [runImpl(st, String(content), {
                        size: normalized.size ?? opts.fontSize,
                        bold,
                        italics: normalized.italics,
                        color: normalized.color,
                        font: normalized.font,
                    })],
            })];
    return new docx_1.TableCell({
        borders: opts.borders ?? st.BORDERS_ALL,
        width: { size: width, type: docx_1.WidthType.DXA },
        verticalAlign: docx_1.VerticalAlign.CENTER,
        margins: st.TABLE_CELL_MARGINS,
        columnSpan: normalized.colspan,
        rowSpan: normalized.rowspan,
        shading: normalized.shading,
        children,
    });
}
function makeTableImpl(st, rows, opts = {}) {
    const count = rows.reduce((max, row) => Math.max(max, row.length), 0);
    const effectiveStyle = opts.contentWidth ? { ...st, CONTENT_WIDTH: opts.contentWidth } : st;
    const widths = tableColumnWidths(effectiveStyle, count, opts.widths);
    const headerRow = opts.headerRow ?? 0;
    const defaultAlign = opts.align ?? docx_1.AlignmentType.JUSTIFIED;
    const centerHeader = opts.centerHeader !== undefined ? opts.centerHeader : true;
    return new docx_1.Table({
        width: { size: effectiveStyle.CONTENT_WIDTH, type: docx_1.WidthType.DXA },
        columnWidths: widths,
        rows: rows.map((row, rowIndex) => {
            const isHeader = rowIndex === headerRow;
            return new docx_1.TableRow({
                children: row.map((cell, cellIndex) => tableCellImpl(st, cell, widths[cellIndex], {
                    ...opts,
                    align: isHeader && centerHeader ? docx_1.AlignmentType.CENTER : defaultAlign,
                    bold: isHeader,
                    fontSize: opts.fontSize ?? st.SIZE,
                    borders: opts.borders ?? st.BORDERS_ALL,
                })),
            });
        }),
    });
}
function makeFooterCityImpl(st, year) {
    return new docx_1.Footer({ children: [centeredImpl(st, `Таганрог ${year}`, { size: st.SIZE })] });
}
function makeFooterPageNumImpl(st) {
    return new docx_1.Footer({
        children: [
            new docx_1.Paragraph({
                alignment: docx_1.AlignmentType.CENTER,
                children: [new docx_1.TextRun({ children: [docx_1.PageNumber.CURRENT], font: st.FONT, size: st.SIZE })],
            }),
        ],
    });
}
function normalizeSubtitle(subtitle) {
    if (Array.isArray(subtitle))
        return subtitle;
    return subtitle === undefined || subtitle === null ? [] : [subtitle];
}
function makeTitlePageImpl(st, opts) {
    const children = [
        titleLine(st, "Министерство науки и высшего образования Российской Федерации", { size: st.SIZE_HEADER }),
        titleLine(st, "Федеральное государственное автономное образовательное", { size: st.SIZE_HEADER }),
        titleLine(st, "учреждение высшего образования", { size: st.SIZE_HEADER }),
        titleLine(st, "«ЮЖНЫЙ ФЕДЕРАЛЬНЫЙ УНИВЕРСИТЕТ»", { size: st.SIZE_HEADER, bold: true }),
        titleLine(st, "Инженерно-технологическая академия", { size: st.SIZE_HEADER }),
        titleLine(st, "Институт компьютерных технологий и информационной безопасности", { size: st.SIZE_HEADER }),
        blankImpl(st), blankImpl(st), blankImpl(st),
        titleLine(st, opts.workType, { size: st.SIZE_TITLE, bold: true }),
        ...normalizeSubtitle(opts.subtitle).map((line) => titleLine(st, line, { size: st.SIZE })),
        titleLine(st, opts.discipline, { size: st.SIZE, after: 120 }),
        blankImpl(st), blankImpl(st), blankImpl(st), blankImpl(st), blankImpl(st), blankImpl(st), blankImpl(st), blankImpl(st), blankImpl(st),
        new docx_1.Paragraph({ spacing: { line: st.LINE }, indent: { left: st.INDENT }, children: [runImpl(st, "Выполнил")] }),
        new docx_1.Paragraph({
            spacing: { line: st.LINE },
            tabStops: st.TITLE_TAB_STOPS,
            children: [runImpl(st, `студент группы ${opts.group}`), new docx_1.Tab(), runImpl(st, "_______________"), new docx_1.Tab(), runImpl(st, opts.author)],
        }),
        blankImpl(st),
        new docx_1.Paragraph({ spacing: { line: st.LINE }, indent: { left: st.INDENT }, children: [runImpl(st, "Принял")] }),
        new docx_1.Paragraph({
            spacing: { line: st.LINE, after: 560 },
            tabStops: st.TITLE_TAB_STOPS,
            children: [runImpl(st, opts.teacherTitle ?? "Доцент кафедры"), new docx_1.Tab(), runImpl(st, "_______________"), new docx_1.Tab(), runImpl(st, opts.teacher)],
        }),
    ];
    return { properties: { page: st.PAGE }, footers: { default: makeFooterCityImpl(st, opts.year) }, children };
}
function makeContentSectionImpl(st, children) {
    return { properties: { page: st.PAGE }, footers: { default: makeFooterPageNumImpl(st) }, children };
}
function makeDocumentImpl(st, sections, opts = {}) {
    return new docx_1.Document({
        settings: { updateFields: true },
        styles: opts.styles ?? defaultStyles(st),
        sections,
    });
}
function saveDocumentImpl(doc, outputPath) {
    return docx_1.Packer.toBuffer(doc).then((buffer) => fs_1.default.promises.writeFile(outputPath, buffer));
}
function createDocxGost(factoryOpts = {}) {
    const st = createDocxStyle(factoryOpts.style);
    return {
        style: st,
        run: (text, opts = {}) => runImpl(st, text, opts),
        paragraph: (content, opts = {}) => paragraphImpl(st, content, opts),
        centered: (content, opts = {}) => centeredImpl(st, content, opts),
        paragraphWithMath: (parts, opts = {}) => paragraphWithMathImpl(st, parts, opts),
        h1: (text, opts = {}) => h1Impl(st, text, opts),
        h2: (text, opts = {}) => h2Impl(st, text, opts),
        h3: (text, opts = {}) => h3Impl(st, text, opts),
        blank: (opts = {}) => blankImpl(st, opts),
        pageBreak: () => pageBreakImpl(),
        caption: (text, opts = {}) => captionImpl(st, text, opts),
        tableCaption: (text, opts = {}) => tableCaptionImpl(st, text, opts),
        createCaptionCounter: (prefix, opts = {}) => createCaptionCounterImpl(st, prefix, opts),
        placeholder: (label, opts = {}) => placeholderImpl(st, label, opts),
        imageBlock: (src, w, h, opts = {}) => imageBlockImpl(st, src, w, h, opts),
        codeLine: (text, opts = {}) => codeLineImpl(st, text, opts),
        codeBlock: (lines, opts = {}) => lines.map((line) => codeLineImpl(st, line, opts)),
        formulaMath: (content, opts = {}) => formulaMathImpl(st, content, opts),
        formulaInline: (label, math, opts = {}) => formulaInlineImpl(st, label, math, opts),
        makeTable: (rows, opts = {}) => makeTableImpl(st, rows, opts),
        makeTitlePage: (opts) => makeTitlePageImpl(st, opts),
        makeContentSection: (children) => makeContentSectionImpl(st, children),
        makeDocument: (sections, opts = {}) => makeDocumentImpl(st, sections, opts),
        saveDocument: (doc, outputPath) => saveDocumentImpl(doc, outputPath),
    };
}
//# sourceMappingURL=docx-gost.js.map
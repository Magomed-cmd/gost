"use strict";

const fs = require("fs");
const {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  HeadingLevel,
  ImageRun,
  Math: DocxMath,
  MathCurlyBrackets,
  MathFraction,
  MathIntegral,
  MathRadical,
  MathRoundBrackets,
  MathRun,
  MathSquareBrackets,
  MathSubScript,
  MathSubSuperScript,
  MathSum,
  MathSuperScript,
  Packer,
  PageBreak,
  PageNumber,
  Paragraph,
  Tab,
  Table,
  TableCell,
  TableOfContents,
  TableRow,
  TabStopType,
  TextRun,
  VerticalAlign,
  WidthType,
} = require("docx");

// ===================== CONSTANTS =====================
const FONT = "Times New Roman";
const SIZE = 28;
const SIZE_TITLE = 32;
const SIZE_HEADER = 24;
const SIZE_CODE = 24;
const INDENT = 709;
const LINE = 360;
const CONTENT_WIDTH = 9355;

const PAGE = {
  size: { width: 11906, height: 16838 },
  margin: { top: 1134, bottom: 1134, left: 1701, right: 850 },
};

const BORDER_SINGLE = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
const BORDERS_ALL = {
  top: BORDER_SINGLE,
  bottom: BORDER_SINGLE,
  left: BORDER_SINGLE,
  right: BORDER_SINGLE,
};

const TITLE_TAB_STOPS = [
  { type: TabStopType.LEFT, position: 4500 },
  { type: TabStopType.RIGHT, position: 9355 },
];

const DEFAULT_DOCX_STYLE = {
  FONT,
  SIZE,
  SIZE_TITLE,
  SIZE_HEADER,
  SIZE_CODE,
  INDENT,
  LINE,
  CONTENT_WIDTH,
  PAGE,
  BORDER_SINGLE,
  BORDERS_ALL,
  TITLE_TAB_STOPS,
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
};

/**
 * Creates a configurable DOCX style object based on GOST defaults.
 *
 * @param {object} [opts={}] Style override options.
 * @returns {object} DOCX style configuration.
 */
function createDocxStyle(opts = {}) {
  const page = opts.page ?? opts.PAGE ?? PAGE;
  const borderSingle = opts.borderSingle ?? opts.BORDER_SINGLE ?? BORDER_SINGLE;
  const bordersAll = opts.bordersAll ?? opts.BORDERS_ALL ?? {
    top: borderSingle,
    bottom: borderSingle,
    left: borderSingle,
    right: borderSingle,
  };

  return {
    FONT: opts.font ?? opts.FONT ?? FONT,
    SIZE: opts.size ?? opts.SIZE ?? SIZE,
    SIZE_TITLE: opts.sizeTitle ?? opts.SIZE_TITLE ?? SIZE_TITLE,
    SIZE_HEADER: opts.sizeHeader ?? opts.SIZE_HEADER ?? SIZE_HEADER,
    SIZE_CODE: opts.sizeCode ?? opts.SIZE_CODE ?? SIZE_CODE,
    INDENT: opts.indent ?? opts.INDENT ?? INDENT,
    LINE: opts.line ?? opts.LINE ?? LINE,
    CONTENT_WIDTH: opts.contentWidth ?? opts.CONTENT_WIDTH ?? CONTENT_WIDTH,
    PAGE: {
      size: { ...(PAGE.size), ...(page.size ?? {}) },
      margin: { ...(PAGE.margin), ...(page.margin ?? {}) },
    },
    BORDER_SINGLE: borderSingle,
    BORDERS_ALL: bordersAll,
    TITLE_TAB_STOPS: opts.titleTabStops ?? opts.TITLE_TAB_STOPS ?? TITLE_TAB_STOPS,
    CODE_FONT: opts.codeFont ?? opts.CODE_FONT ?? DEFAULT_DOCX_STYLE.CODE_FONT,
    TABLE_CELL_LINE: opts.tableCellLine ?? opts.TABLE_CELL_LINE ?? DEFAULT_DOCX_STYLE.TABLE_CELL_LINE,
    TABLE_CELL_MARGINS: opts.tableCellMargins ?? opts.TABLE_CELL_MARGINS ?? DEFAULT_DOCX_STYLE.TABLE_CELL_MARGINS,
    HEADING_COLOR: opts.headingColor ?? opts.HEADING_COLOR ?? DEFAULT_DOCX_STYLE.HEADING_COLOR,
    CAPTION_BEFORE: opts.captionBefore ?? opts.CAPTION_BEFORE ?? DEFAULT_DOCX_STYLE.CAPTION_BEFORE,
    CAPTION_AFTER: opts.captionAfter ?? opts.CAPTION_AFTER ?? DEFAULT_DOCX_STYLE.CAPTION_AFTER,
    TABLE_CAPTION_BEFORE: opts.tableCaptionBefore ?? opts.TABLE_CAPTION_BEFORE ?? DEFAULT_DOCX_STYLE.TABLE_CAPTION_BEFORE,
    TABLE_CAPTION_AFTER: opts.tableCaptionAfter ?? opts.TABLE_CAPTION_AFTER ?? DEFAULT_DOCX_STYLE.TABLE_CAPTION_AFTER,
    IMAGE_BEFORE: opts.imageBefore ?? opts.IMAGE_BEFORE ?? DEFAULT_DOCX_STYLE.IMAGE_BEFORE,
    IMAGE_AFTER: opts.imageAfter ?? opts.IMAGE_AFTER ?? DEFAULT_DOCX_STYLE.IMAGE_AFTER,
    FORMULA_BEFORE: opts.formulaBefore ?? opts.FORMULA_BEFORE ?? DEFAULT_DOCX_STYLE.FORMULA_BEFORE,
    FORMULA_AFTER: opts.formulaAfter ?? opts.FORMULA_AFTER ?? DEFAULT_DOCX_STYLE.FORMULA_AFTER,
  };
}

function styleFromOpts(opts = {}) {
  return opts.style ? createDocxStyle(opts.style) : DEFAULT_DOCX_STYLE;
}

// ===================== INTERNAL HELPERS =====================
function isMathLike(value) {
  return value instanceof DocxMath
    || (
      value
      && typeof value === "object"
      && typeof value.constructor?.name === "string"
      && value.constructor.name.startsWith("Math")
    );
}

function paragraphChildren(content, opts = {}) {
  if (typeof content === "string" || typeof content === "number") {
    return [run(String(content), opts)];
  }

  return content.map((item) => {
    if (typeof item === "string" || typeof item === "number") {
      return run(String(item), opts);
    }

    return run(item.text, {
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

function paragraphOptions(opts = {}, children = []) {
  const style = styleFromOpts(opts);
  return {
    heading: opts.heading,
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing: {
      line: opts.line ?? style.LINE,
      before: opts.before ?? 0,
      after: opts.after ?? 0,
    },
    indent: opts.noIndent ? undefined : { firstLine: opts.indent ?? style.INDENT },
    keepNext: !!opts.keepNext,
    keepLines: opts.keepLines !== undefined ? opts.keepLines : true,
    children,
  };
}

function titleLine(content, opts = {}) {
  const style = styleFromOpts(opts);
  return centered(content, {
    ...opts,
    size: opts.size ?? style.SIZE,
    line: opts.line ?? style.LINE,
    before: opts.before ?? 0,
    after: opts.after ?? 0,
  });
}

function normalizeSubtitle(subtitle) {
  if (Array.isArray(subtitle)) {
    return subtitle;
  }

  return subtitle === undefined || subtitle === null ? [] : [subtitle];
}

function defaultStyles(style = DEFAULT_DOCX_STYLE) {
  return {
    default: {
      document: {
        run: { font: style.FONT, size: style.SIZE },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        run: { font: style.FONT, size: style.SIZE, bold: true, color: style.HEADING_COLOR },
        paragraph: {
          alignment: AlignmentType.CENTER,
          spacing: { before: 240, after: 120 },
        },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        run: { font: style.FONT, size: style.SIZE, bold: true, color: style.HEADING_COLOR },
        paragraph: { spacing: { before: 200, after: 80 } },
      },
      {
        id: "Heading3",
        name: "Heading 3",
        basedOn: "Normal",
        next: "Normal",
        run: { font: style.FONT, size: style.SIZE, bold: true, italics: true, color: style.HEADING_COLOR },
        paragraph: { spacing: { before: 140, after: 60 } },
      },
    ],
  };
}

function tableColumnWidths(columnCount, widths, style = DEFAULT_DOCX_STYLE) {
  if (Array.isArray(widths) && widths.length > 0) {
    return widths;
  }

  const baseWidth = Math.floor(style.CONTENT_WIDTH / columnCount);
  const result = Array(columnCount).fill(baseWidth);
  result[result.length - 1] += style.CONTENT_WIDTH - result.reduce((sum, width) => sum + width, 0);
  return result;
}

function normalizeCell(cell) {
  if (cell && typeof cell === "object" && !Array.isArray(cell) && !isMathLike(cell)) {
    return cell;
  }

  return { text: cell };
}

function tableCell(cell, width, opts = {}) {
  const style = styleFromOpts(opts);
  const normalized = normalizeCell(cell);
  const content = normalized.text ?? "";
  const align = normalized.align || opts.align;
  const textBold = normalized.bold !== undefined ? normalized.bold : opts.bold;
  const children = Array.isArray(content) || isMathLike(content)
    ? [paragraphWithMath(Array.isArray(content) ? content : [content], {
      align,
      noIndent: true,
      line: opts.cellLine ?? style.TABLE_CELL_LINE,
      size: opts.fontSize,
    })]
    : [new Paragraph({
      alignment: align,
      spacing: { line: opts.cellLine ?? style.TABLE_CELL_LINE, before: 0, after: 0 },
      children: [run(String(content), {
        size: normalized.size ?? opts.fontSize,
        bold: textBold,
        italics: normalized.italics,
        color: normalized.color,
        font: normalized.font,
      })],
    })];

  return new TableCell({
    borders: opts.borders,
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    margins: opts.cellMargins ?? style.TABLE_CELL_MARGINS,
    columnSpan: normalized.colspan,
    rowSpan: normalized.rowspan,
    shading: normalized.shading,
    children,
  });
}

// ===================== TEXT PRIMITIVES =====================

/**
 * Creates a Times New Roman text run with GOST defaults.
 *
 * @param {string|number} text Text content.
 * @param {object} [opts={}] Text run options.
 * @returns {TextRun} Configured docx text run.
 */
function run(text, opts = {}) {
  const style = styleFromOpts(opts);
  return new TextRun({
    text: String(text ?? ""),
    font: opts.font || style.FONT,
    size: opts.size ?? style.SIZE,
    bold: !!opts.bold,
    italics: !!opts.italics,
    allCaps: !!opts.allCaps,
    underline: opts.underline ? (opts.underline === true ? {} : opts.underline) : undefined,
    color: opts.color,
  });
}

/**
 * Creates a paragraph from text or text-run descriptors.
 *
 * @param {string|number|Array<object|string|number>} content Paragraph content.
 * @param {object} [opts={}] Paragraph options.
 * @returns {Paragraph} Configured docx paragraph.
 */
function paragraph(content, opts = {}) {
  return new Paragraph(paragraphOptions(opts, paragraphChildren(content, opts)));
}

/**
 * Creates a centered paragraph without first-line indent.
 *
 * @param {string|number|Array<object|string|number>} content Paragraph content.
 * @param {object} [opts={}] Paragraph options.
 * @returns {Paragraph} Configured centered paragraph.
 */
function centered(content, opts = {}) {
  return paragraph(content, { ...opts, align: AlignmentType.CENTER, noIndent: true });
}

// ===================== HEADINGS =====================

/**
 * Creates a first-level centered heading.
 *
 * @param {string} text Heading text.
 * @returns {Paragraph} Heading paragraph.
 */
function h1(text, opts = {}) {
  const style = styleFromOpts(opts);
  return paragraph([{ text: String(text).toUpperCase(), bold: true, color: style.HEADING_COLOR }], {
    ...opts,
    heading: HeadingLevel.HEADING_1,
    align: AlignmentType.CENTER,
    noIndent: true,
    before: 240,
    after: 120,
    keepNext: true,
  });
}

/**
 * Creates a second-level heading.
 *
 * @param {string} text Heading text.
 * @returns {Paragraph} Heading paragraph.
 */
function h2(text, opts = {}) {
  const style = styleFromOpts(opts);
  return paragraph([{ text, bold: true, color: style.HEADING_COLOR }], {
    ...opts,
    heading: HeadingLevel.HEADING_2,
    before: 200,
    after: 80,
    keepNext: true,
  });
}

/**
 * Creates a third-level heading.
 *
 * @param {string} text Heading text.
 * @returns {Paragraph} Heading paragraph.
 */
function h3(text, opts = {}) {
  const style = styleFromOpts(opts);
  return paragraph([{ text, bold: true, italics: true, color: style.HEADING_COLOR }], {
    ...opts,
    heading: HeadingLevel.HEADING_3,
    before: 140,
    after: 60,
    keepNext: true,
  });
}

// ===================== STRUCTURAL ELEMENTS =====================

/**
 * Creates an empty paragraph line.
 *
 * @returns {Paragraph} Empty paragraph.
 */
function blank(opts = {}) {
  const style = styleFromOpts(opts);
  return new Paragraph({ spacing: { line: opts.line ?? style.LINE }, children: [run("", opts)] });
}

/**
 * Creates a paragraph containing a page break.
 *
 * @returns {Paragraph} Page break paragraph.
 */
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

/**
 * Creates a centered figure or table caption.
 *
 * @param {string} text Caption text.
 * @returns {Paragraph} Caption paragraph.
 */
function caption(text, opts = {}) {
  const style = styleFromOpts(opts);
  return centered(text, {
    ...opts,
    before: opts.before ?? style.CAPTION_BEFORE,
    after: opts.after ?? style.CAPTION_AFTER,
    size: opts.size ?? style.SIZE,
  });
}

/**
 * Table caption per ГОСТ 7.32: left-aligned, no indent, placed ABOVE the table.
 * Format: "Таблица N – Название"
 *
 * @param {string} text Full caption text including prefix and number.
 * @returns {Paragraph} Left-aligned caption paragraph.
 */
function tableCaption(text, opts = {}) {
  const style = styleFromOpts(opts);
  return paragraph(text, {
    ...opts,
    align: AlignmentType.LEFT,
    noIndent: true,
    before: opts.before ?? style.TABLE_CAPTION_BEFORE,
    after: opts.after ?? style.TABLE_CAPTION_AFTER,
    size: opts.size ?? style.SIZE,
  });
}

/**
 * Creates a per-document numbered caption counter.
 *
 * @param {string} prefix Caption prefix, for example "Рисунок" or "Таблица".
 * @param {object} [opts={}] Counter options.
 * @param {number} [opts.start=1] First caption number.
 * @param {boolean} [opts.table=false] Use left-aligned tableCaption (ГОСТ) instead of centered caption.
 * @returns {{peek: function(): number, caption: function(string): Paragraph}} Caption counter.
 */
function createCaptionCounter(prefix, opts = {}) {
  let n = opts.start ?? 1;
  const captionFn = opts.table ? tableCaption : caption;

  return {
    peek: () => n,
    caption: (text) => captionFn(`${prefix} ${n++} – ${text}`, opts),
  };
}

/**
 * Creates a bordered placeholder paragraph for a missing image.
 *
 * @param {string} label Placeholder label.
 * @returns {Paragraph} Placeholder paragraph.
 */
function placeholder(label, opts = {}) {
  const border = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" };

  return new Paragraph({
    ...paragraphOptions({
      ...opts,
      align: AlignmentType.CENTER,
      noIndent: true,
      before: 120,
      after: 60,
    }, [run(label, { ...opts, italics: true, color: opts.color ?? "888888" })]),
    border: { top: border, bottom: border, left: border, right: border },
  });
}

/**
 * Creates a centered PNG image paragraph.
 *
 * @param {string|Buffer} imagePath Path to a PNG image or PNG data buffer.
 * @param {number} width Image width in pixels.
 * @param {number} height Image height in pixels.
 * @param {object} [opts={}] Paragraph options.
 * @returns {Paragraph} Image paragraph.
 */
function imageBlock(imagePath, width, height, opts = {}) {
  const style = styleFromOpts(opts);
  const data = Buffer.isBuffer(imagePath) ? imagePath : fs.readFileSync(imagePath);

  return new Paragraph({
    alignment: opts.align || AlignmentType.CENTER,
    spacing: {
      line: opts.line ?? style.LINE,
      before: opts.before ?? style.IMAGE_BEFORE,
      after: opts.after ?? style.IMAGE_AFTER,
    },
    keepNext: !!opts.keepNext,
    keepLines: opts.keepLines !== undefined ? opts.keepLines : true,
    children: [
      new ImageRun({
        data,
        type: "png",
        transformation: { width, height },
      }),
    ],
  });
}

// ===================== CODE LISTINGS =====================

/**
 * Creates one code listing line.
 *
 * @param {string} text Code line text.
 * @param {object} [opts={}] Code line options.
 * @returns {Paragraph} Code paragraph.
 */
function codeLine(text, opts = {}) {
  const style = styleFromOpts(opts);
  return paragraph([{ text, font: opts.font ?? style.CODE_FONT, size: opts.size ?? style.SIZE_CODE }], {
    style: opts.style,
    noIndent: true,
    line: opts.line ?? style.LINE,
  });
}

/**
 * Creates a code block from lines.
 *
 * @param {string[]} lines Code lines.
 * @param {object} [opts={}] Code block options.
 * @returns {Paragraph[]} Code paragraphs.
 */
function codeBlock(lines, opts = {}) {
  return lines.map((line) => codeLine(line, opts));
}

// ===================== MATH HELPERS =====================

/**
 * Converts strings, numbers, arrays and docx math objects to math components.
 *
 * @param {*} value Math component input.
 * @returns {Array} Flat array of docx math components.
 */
function toMathComponents(value) {
  if (value === undefined || value === null) {
    return [];
  }

  if (Array.isArray(value)) {
    return value.flatMap((item) => toMathComponents(item));
  }

  if (typeof value === "string" || typeof value === "number") {
    return [new MathRun(String(value))];
  }

  return [value];
}

/**
 * Creates an inline docx math object.
 *
 * @param {...*} parts Math expression parts.
 * @returns {DocxMath} Inline math object.
 */
function mathExpr(...parts) {
  return new DocxMath({
    children: parts.flatMap((part) => toMathComponents(part)),
  });
}

/**
 * Creates a math fraction.
 *
 * @param {*} numerator Fraction numerator.
 * @param {*} denominator Fraction denominator.
 * @returns {MathFraction} Math fraction.
 */
function mathFraction(numerator, denominator) {
  return new MathFraction({
    numerator: toMathComponents(numerator),
    denominator: toMathComponents(denominator),
  });
}

/**
 * Creates a math subscript.
 *
 * @param {*} base Base expression.
 * @param {*} subScript Subscript expression.
 * @returns {MathSubScript} Math subscript.
 */
function mathSub(base, subScript) {
  return new MathSubScript({
    children: toMathComponents(base),
    subScript: toMathComponents(subScript),
  });
}

/**
 * Creates a math superscript.
 *
 * @param {*} base Base expression.
 * @param {*} superScript Superscript expression.
 * @returns {MathSuperScript} Math superscript.
 */
function mathSup(base, superScript) {
  return new MathSuperScript({
    children: toMathComponents(base),
    superScript: toMathComponents(superScript),
  });
}

/**
 * Creates a math expression with subscript and superscript.
 *
 * @param {*} base Base expression.
 * @param {*} subScript Subscript expression.
 * @param {*} superScript Superscript expression.
 * @returns {MathSubSuperScript} Math subscript/superscript.
 */
function mathSubSup(base, subScript, superScript) {
  return new MathSubSuperScript({
    children: toMathComponents(base),
    subScript: toMathComponents(subScript),
    superScript: toMathComponents(superScript),
  });
}

/**
 * Wraps math parts in round brackets.
 *
 * @param {...*} parts Math expression parts.
 * @returns {MathRoundBrackets} Round-bracket math object.
 */
function mathParen(...parts) {
  return new MathRoundBrackets({
    children: parts.flatMap((part) => toMathComponents(part)),
  });
}

/**
 * Wraps math parts in square brackets.
 *
 * @param {...*} parts Math expression parts.
 * @returns {MathSquareBrackets} Square-bracket math object.
 */
function mathBracket(...parts) {
  return new MathSquareBrackets({
    children: parts.flatMap((part) => toMathComponents(part)),
  });
}

/**
 * Wraps math parts in curly brackets.
 *
 * @param {...*} parts Math expression parts.
 * @returns {MathCurlyBrackets} Curly-bracket math object.
 */
function mathBrace(...parts) {
  return new MathCurlyBrackets({
    children: parts.flatMap((part) => toMathComponents(part)),
  });
}

/**
 * Creates a math radical.
 *
 * @param {*} content Radical content.
 * @param {*} [degree] Optional radical degree.
 * @returns {MathRadical} Math radical.
 */
function mathRoot(content, degree) {
  return new MathRadical({
    children: toMathComponents(content),
    degree: degree === undefined ? undefined : toMathComponents(degree),
  });
}

/**
 * Creates a math sum.
 *
 * @param {*} content Sum content.
 * @param {object} [opts={}] Sum limits options.
 * @returns {MathSum} Math sum.
 */
function mathSum(content, opts = {}) {
  return new MathSum({
    children: toMathComponents(content),
    subScript: opts.subScript ? toMathComponents(opts.subScript) : undefined,
    superScript: opts.superScript ? toMathComponents(opts.superScript) : undefined,
  });
}

/**
 * Creates a math integral.
 *
 * @param {*} content Integral content.
 * @param {object} [opts={}] Integral limits options.
 * @returns {MathIntegral} Math integral.
 */
function mathIntegral(content, opts = {}) {
  return new MathIntegral({
    children: toMathComponents(content),
    subScript: opts.subScript ? toMathComponents(opts.subScript) : undefined,
    superScript: opts.superScript ? toMathComponents(opts.superScript) : undefined,
  });
}

/**
 * Creates a centered standalone formula paragraph.
 *
 * @param {*} content Formula content.
 * @param {object} [opts={}] Paragraph options.
 * @returns {Paragraph} Formula paragraph.
 */
function formulaMath(content, opts = {}) {
  const style = styleFromOpts(opts);
  return new Paragraph(paragraphOptions({
    ...opts,
    align: AlignmentType.CENTER,
    noIndent: true,
    before: opts.before ?? style.FORMULA_BEFORE,
    after: opts.after ?? style.FORMULA_AFTER,
  }, [mathExpr(content)]));
}

/**
 * Creates a paragraph with a text label followed by inline math.
 *
 * @param {string} label Text label.
 * @param {*} mathContent Math content.
 * @param {object} [opts={}] Paragraph options.
 * @returns {Paragraph} Paragraph with inline formula.
 */
function formulaInline(label, mathContent, opts = {}) {
  return new Paragraph(paragraphOptions(opts, [
    run(`${label} `, opts),
    mathExpr(mathContent),
  ]));
}

/**
 * Creates a paragraph from mixed text and docx math parts.
 *
 * @param {Array} parts Mixed text, descriptor and math parts.
 * @param {object} [opts={}] Paragraph options.
 * @returns {Paragraph} Paragraph with mixed content.
 */
function paragraphWithMath(parts, opts = {}) {
  const children = parts.map((part) => {
    if (typeof part === "string" || typeof part === "number") {
      return run(String(part), opts);
    }

    if (part instanceof DocxMath) {
      return part;
    }

    if (Array.isArray(part) || isMathLike(part)) {
      return mathExpr(part);
    }

    return run(part.text, {
      ...opts,
      size: part.size ?? opts.size,
      bold: part.bold ?? opts.bold,
      italics: part.italics ?? opts.italics,
      allCaps: part.allCaps ?? opts.allCaps,
      underline: part.underline ?? opts.underline,
      color: part.color ?? opts.color,
      font: part.font ?? opts.font,
    });
  });

  return new Paragraph(paragraphOptions(opts, children));
}

// ===================== TABLES =====================

/**
 * Creates a GOST-styled table.
 *
 * @param {Array<Array<string|number|object>>} rows Table rows.
 * @param {object} [opts={}] Table options.
 * @returns {Table} Configured docx table.
 */
function makeTable(rows, opts = {}) {
  const style = styleFromOpts(opts);
  const columnCount = rows.reduce((max, row) => Math.max(max, row.length), 0);
  const tableWidth = opts.contentWidth ?? style.CONTENT_WIDTH;
  const tableStyle = opts.contentWidth ? { ...style, CONTENT_WIDTH: tableWidth } : style;
  const widths = tableColumnWidths(columnCount, opts.widths, tableStyle);
  const headerRow = opts.headerRow ?? 0;
  const borders = opts.borders || style.BORDERS_ALL;
  const fontSize = opts.fontSize ?? style.SIZE;
  const defaultAlign = opts.align || AlignmentType.JUSTIFIED;
  const centerHeader = opts.centerHeader !== undefined ? opts.centerHeader : true;

  return new Table({
    width: { size: tableWidth, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((row, rowIndex) => {
      const isHeader = rowIndex === headerRow;

      return new TableRow({
        children: row.map((cell, cellIndex) => tableCell(cell, widths[cellIndex], {
          align: isHeader && centerHeader ? AlignmentType.CENTER : defaultAlign,
          bold: isHeader,
          borders,
          fontSize,
          style: opts.style,
          cellLine: opts.cellLine,
          cellMargins: opts.cellMargins,
        })),
      });
    }),
  });
}

// ===================== DOCUMENT STRUCTURE =====================

/**
 * Creates a city/year footer for the title page.
 *
 * @param {string|number} year Footer year.
 * @returns {Footer} Configured footer.
 */
function makeFooterCity(year, opts = {}) {
  const style = styleFromOpts(opts);
  return new Footer({
    children: [
      centered(`Таганрог ${year}`, { ...opts, size: opts.size ?? style.SIZE }),
    ],
  });
}

/**
 * Creates a centered page-number footer.
 *
 * @returns {Footer} Configured footer.
 */
function makeFooterPageNum(opts = {}) {
  const style = styleFromOpts(opts);
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ children: [PageNumber.CURRENT], font: style.FONT, size: opts.size ?? style.SIZE }),
        ],
      }),
    ],
  });
}

/**
 * Creates a title-page section.
 *
 * @param {object} opts Title-page options.
 * @returns {object} Docx document section.
 */
function makeTitlePage(opts) {
  const style = styleFromOpts(opts);
  const children = [
    titleLine("Министерство науки и высшего образования Российской Федерации", { style: opts.style, size: style.SIZE_HEADER }),
    titleLine("Федеральное государственное автономное образовательное", { style: opts.style, size: style.SIZE_HEADER }),
    titleLine("учреждение высшего образования", { style: opts.style, size: style.SIZE_HEADER }),
    titleLine("«ЮЖНЫЙ ФЕДЕРАЛЬНЫЙ УНИВЕРСИТЕТ»", { style: opts.style, size: style.SIZE_HEADER, bold: true }),
    titleLine("Инженерно-технологическая академия", { style: opts.style, size: style.SIZE_HEADER }),
    titleLine("Институт компьютерных технологий и информационной безопасности", { style: opts.style, size: style.SIZE_HEADER }),
    blank({ style: opts.style }),
    blank({ style: opts.style }),
    blank({ style: opts.style }),
    titleLine(opts.workType, { style: opts.style, size: style.SIZE_TITLE, bold: true }),
    ...normalizeSubtitle(opts.subtitle).map((line) => titleLine(line, { style: opts.style, size: style.SIZE })),
    titleLine(opts.discipline, { style: opts.style, size: style.SIZE, after: 120 }),
    blank({ style: opts.style }),
    blank({ style: opts.style }),
    blank({ style: opts.style }),
    blank({ style: opts.style }),
    blank({ style: opts.style }),
    blank({ style: opts.style }),
    blank({ style: opts.style }),
    blank({ style: opts.style }),
    blank({ style: opts.style }),
    new Paragraph({
      spacing: { line: style.LINE },
      indent: { left: style.INDENT },
      children: [run("Выполнил", { style: opts.style })],
    }),
    new Paragraph({
      spacing: { line: style.LINE },
      tabStops: style.TITLE_TAB_STOPS,
      children: [
        run(`студент группы ${opts.group}`, { style: opts.style }),
        new Tab(),
        run("_______________", { style: opts.style }),
        new Tab(),
        run(opts.author, { style: opts.style }),
      ],
    }),
    blank({ style: opts.style }),
    new Paragraph({
      spacing: { line: style.LINE },
      indent: { left: style.INDENT },
      children: [run("Принял", { style: opts.style })],
    }),
    new Paragraph({
      spacing: { line: style.LINE, after: 560 },
      tabStops: style.TITLE_TAB_STOPS,
      children: [
        run(opts.teacherTitle ?? "Доцент кафедры", { style: opts.style }),
        new Tab(),
        run("_______________", { style: opts.style }),
        new Tab(),
        run(opts.teacher, { style: opts.style }),
      ],
    }),
  ];

  return {
    properties: { page: style.PAGE },
    footers: { default: makeFooterCity(opts.year, { style: opts.style }) },
    children,
  };
}

/**
 * Creates a content section with page numbering.
 *
 * @param {Array} children Section children.
 * @param {string|number} [year] Reserved for compatibility.
 * @returns {object} Docx document section.
 */
function makeContentSection(children, year, opts = {}) {
  if (year && typeof year === "object") {
    opts = year;
  }
  const style = styleFromOpts(opts);

  return {
    properties: { page: style.PAGE },
    footers: { default: makeFooterPageNum({ style: opts.style }) },
    children,
  };
}

/**
 * Creates a docx document with GOST default styles.
 *
 * @param {Array<object>} sections Document sections.
 * @param {object} [opts={}] Document options.
 * @returns {Document} Configured docx document.
 */
function makeDocument(sections, opts = {}) {
  const style = styleFromOpts(opts);
  return new Document({
    settings: {
      updateFields: true,
    },
    styles: opts.styles || defaultStyles(style),
    sections,
  });
}

/**
 * Saves a docx document to disk.
 *
 * @param {Document} doc Docx document.
 * @param {string} outputPath Output file path.
 * @returns {Promise<void>} Promise resolved after writing the file.
 */
function saveDocument(doc, outputPath) {
  return Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(outputPath, buffer);
  });
}

module.exports = {
  FONT,
  SIZE,
  SIZE_TITLE,
  SIZE_HEADER,
  SIZE_CODE,
  INDENT,
  LINE,
  CONTENT_WIDTH,
  PAGE,
  BORDER_SINGLE,
  BORDERS_ALL,
  TITLE_TAB_STOPS,
  DEFAULT_DOCX_STYLE,
  createDocxStyle,

  run,
  paragraph,
  centered,
  paragraphWithMath,

  h1,
  h2,
  h3,

  blank,
  pageBreak,
  caption,
  tableCaption,
  createCaptionCounter,
  placeholder,
  imageBlock,

  codeLine,
  codeBlock,

  toMathComponents,
  mathExpr,
  mathFraction,
  mathSub,
  mathSup,
  mathSubSup,
  mathParen,
  mathBracket,
  mathBrace,
  mathRoot,
  mathSum,
  mathIntegral,
  formulaMath,
  formulaInline,

  makeTable,

  makeFooterCity,
  makeFooterPageNum,
  makeTitlePage,
  makeContentSection,
  makeDocument,
  saveDocument,

  TableOfContents,
};

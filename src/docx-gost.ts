import fs from "fs";
import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  HeadingLevel,
  ImageRun,
  Math as DocxMath,
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
} from "docx";

type DocxBorder = { style: (typeof BorderStyle)[keyof typeof BorderStyle]; size: number; color: string };
type DocxBorders = { top: DocxBorder; bottom: DocxBorder; left: DocxBorder; right: DocxBorder };
type TabStop = { type: (typeof TabStopType)[keyof typeof TabStopType]; position: number };
// The docx package exposes many separate math component classes with narrow
// constructor types; using any here keeps wrappers compatible with all of them.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type MathComponent = any;
type ParagraphChild = TextRun | DocxMath;

export interface DocxStyleConfig {
  FONT: string;
  SIZE: number;
  SIZE_TITLE: number;
  SIZE_HEADER: number;
  SIZE_CODE: number;
  INDENT: number;
  LINE: number;
  CONTENT_WIDTH: number;
  CODE_FONT: string;
  TABLE_CELL_LINE: number;
  TABLE_CELL_MARGINS: { top: number; bottom: number; left: number; right: number };
  HEADING_COLOR: string;
  CAPTION_BEFORE: number;
  CAPTION_AFTER: number;
  TABLE_CAPTION_BEFORE: number;
  TABLE_CAPTION_AFTER: number;
  IMAGE_BEFORE: number;
  IMAGE_AFTER: number;
  FORMULA_BEFORE: number;
  FORMULA_AFTER: number;
  PAGE: {
    size: { width: number; height: number };
    margin: { top: number; bottom: number; left: number; right: number };
  };
  BORDER_SINGLE: DocxBorder;
  BORDERS_ALL: DocxBorders;
  TITLE_TAB_STOPS: TabStop[];
}

export interface DocxGostOptions {
  style?: Partial<DocxStyleConfig> & {
    font?: string;
    size?: number;
    sizeTitle?: number;
    sizeHeader?: number;
    sizeCode?: number;
    indent?: number;
    line?: number;
    contentWidth?: number;
    codeFont?: string;
    page?: Partial<DocxStyleConfig["PAGE"]>;
  };
}

export interface RunDescriptor {
  text: string;
  size?: number;
  bold?: boolean;
  italics?: boolean;
  allCaps?: boolean;
  underline?: boolean | object;
  color?: string;
  font?: string;
}

export type Content = string | number | Array<string | number | RunDescriptor>;

export interface ParagraphOpts {
  align?: (typeof AlignmentType)[keyof typeof AlignmentType] | string;
  noIndent?: boolean;
  indent?: number;
  line?: number;
  before?: number;
  after?: number;
  bold?: boolean;
  italics?: boolean;
  allCaps?: boolean;
  underline?: boolean | object;
  size?: number;
  color?: string;
  font?: string;
  keepNext?: boolean;
  keepLines?: boolean;
  imageType?: "png" | "jpg" | "gif" | "bmp";
}

export interface TableOpts {
  widths?: number[];
  headerRow?: number;
  align?: (typeof AlignmentType)[keyof typeof AlignmentType] | string;
  centerHeader?: boolean;
  fontSize?: number;
  cellLine?: number;
  contentWidth?: number;
  borders?: DocxBorders;
}

export interface CaptionCounterOpts extends ParagraphOpts {
  start?: number;
  table?: boolean;
}

export interface CaptionCounter {
  peek: () => number;
  caption: (text: string) => Paragraph;
}

export interface CellDescriptor {
  text: CellValue | CellValue[];
  align?: (typeof AlignmentType)[keyof typeof AlignmentType] | string;
  bold?: boolean;
  italics?: boolean;
  size?: number;
  color?: string;
  font?: string;
  colspan?: number;
  rowspan?: number;
  shading?: object;
}

export type CellValue = string | number | CellDescriptor | MathComponent;

export interface TitlePageOpts {
  workType: string;
  subtitle?: string | string[];
  discipline: string;
  group: string;
  author: string;
  teacher: string;
  teacherTitle?: string;
  year: number | string;
}

export interface DocxSection {
  properties: { page: DocxStyleConfig["PAGE"] };
  footers: { default: Footer };
  children: Array<Paragraph | Table>;
}

export interface DocxGostInstance {
  readonly style: DocxStyleConfig;
  run(text: string | number, opts?: ParagraphOpts): TextRun;
  paragraph(content: Content, opts?: ParagraphOpts): Paragraph;
  centered(content: Content, opts?: ParagraphOpts): Paragraph;
  paragraphWithMath(parts: Array<string | number | RunDescriptor | MathComponent>, opts?: ParagraphOpts): Paragraph;
  h1(text: string, opts?: ParagraphOpts): Paragraph;
  h2(text: string, opts?: ParagraphOpts): Paragraph;
  h3(text: string, opts?: ParagraphOpts): Paragraph;
  blank(opts?: ParagraphOpts): Paragraph;
  pageBreak(): Paragraph;
  caption(text: string, opts?: ParagraphOpts): Paragraph;
  tableCaption(text: string, opts?: ParagraphOpts): Paragraph;
  createCaptionCounter(prefix: string, opts?: CaptionCounterOpts): CaptionCounter;
  placeholder(label: string, opts?: ParagraphOpts): Paragraph;
  imageBlock(imagePath: string | Buffer, width: number, height: number, opts?: ParagraphOpts): Paragraph;
  codeLine(text: string, opts?: ParagraphOpts): Paragraph;
  codeBlock(lines: string[], opts?: ParagraphOpts): Paragraph[];
  formulaMath(content: unknown, opts?: ParagraphOpts): Paragraph;
  formulaInline(label: string, mathContent: unknown, opts?: ParagraphOpts): Paragraph;
  makeTable(rows: CellValue[][], opts?: TableOpts): Table;
  makeTitlePage(opts: TitlePageOpts): DocxSection;
  makeContentSection(children: Array<Paragraph | Table>, opts?: object): DocxSection;
  makeDocument(sections: DocxSection[], opts?: { styles?: object }): Document;
  saveDocument(doc: Document, outputPath: string): Promise<void>;
}

export const DEFAULT_DOCX_STYLE: DocxStyleConfig = {
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
  BORDER_SINGLE: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
  BORDERS_ALL: {
    top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
    bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
    left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
    right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
  },
  TITLE_TAB_STOPS: [
    { type: TabStopType.LEFT, position: 4500 },
    { type: TabStopType.RIGHT, position: 9355 },
  ],
};

export function createDocxStyle(opts: DocxGostOptions["style"] = {}): DocxStyleConfig {
  const borderSingle = opts.BORDER_SINGLE ?? DEFAULT_DOCX_STYLE.BORDER_SINGLE;
  return {
    ...DEFAULT_DOCX_STYLE,
    ...opts,
    FONT: opts.font ?? opts.FONT ?? DEFAULT_DOCX_STYLE.FONT,
    SIZE: opts.size ?? opts.SIZE ?? DEFAULT_DOCX_STYLE.SIZE,
    SIZE_TITLE: opts.sizeTitle ?? opts.SIZE_TITLE ?? DEFAULT_DOCX_STYLE.SIZE_TITLE,
    SIZE_HEADER: opts.sizeHeader ?? opts.SIZE_HEADER ?? DEFAULT_DOCX_STYLE.SIZE_HEADER,
    SIZE_CODE: opts.sizeCode ?? opts.SIZE_CODE ?? DEFAULT_DOCX_STYLE.SIZE_CODE,
    INDENT: opts.indent ?? opts.INDENT ?? DEFAULT_DOCX_STYLE.INDENT,
    LINE: opts.line ?? opts.LINE ?? DEFAULT_DOCX_STYLE.LINE,
    CONTENT_WIDTH: opts.contentWidth ?? opts.CONTENT_WIDTH ?? DEFAULT_DOCX_STYLE.CONTENT_WIDTH,
    CODE_FONT: opts.codeFont ?? opts.CODE_FONT ?? DEFAULT_DOCX_STYLE.CODE_FONT,
    PAGE: {
      size: { ...DEFAULT_DOCX_STYLE.PAGE.size, ...(opts.page?.size ?? opts.PAGE?.size ?? {}) },
      margin: { ...DEFAULT_DOCX_STYLE.PAGE.margin, ...(opts.page?.margin ?? opts.PAGE?.margin ?? {}) },
    },
    BORDER_SINGLE: borderSingle,
    BORDERS_ALL: opts.BORDERS_ALL ?? {
      top: borderSingle,
      bottom: borderSingle,
      left: borderSingle,
      right: borderSingle,
    },
    TITLE_TAB_STOPS: opts.TITLE_TAB_STOPS ?? DEFAULT_DOCX_STYLE.TITLE_TAB_STOPS,
  };
}

function isMathLike(value: unknown): value is MathComponent {
  return value instanceof DocxMath
    || (typeof value === "object" && value !== null && value.constructor.name.startsWith("Math"));
}

function runImpl(st: DocxStyleConfig, text: string | number, opts: ParagraphOpts = {}): TextRun {
  return new TextRun({
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

function paragraphOptions(st: DocxStyleConfig, opts: ParagraphOpts = {}, children: ParagraphChild[] = []) {
  return {
    heading: undefined,
    alignment: (opts.align as (typeof AlignmentType)[keyof typeof AlignmentType]) || AlignmentType.JUSTIFIED,
    spacing: { line: opts.line ?? st.LINE, before: opts.before ?? 0, after: opts.after ?? 0 },
    indent: opts.noIndent ? undefined : { firstLine: opts.indent ?? st.INDENT },
    keepNext: !!opts.keepNext,
    keepLines: opts.keepLines !== undefined ? opts.keepLines : true,
    children,
  };
}

function paragraphChildren(st: DocxStyleConfig, content: Content, opts: ParagraphOpts = {}): TextRun[] {
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

function paragraphImpl(st: DocxStyleConfig, content: Content, opts: ParagraphOpts = {}): Paragraph {
  return new Paragraph(paragraphOptions(st, opts, paragraphChildren(st, content, opts)));
}

function centeredImpl(st: DocxStyleConfig, content: Content, opts: ParagraphOpts = {}): Paragraph {
  return paragraphImpl(st, content, { ...opts, align: AlignmentType.CENTER, noIndent: true });
}

function titleLine(st: DocxStyleConfig, content: Content, opts: ParagraphOpts = {}): Paragraph {
  return centeredImpl(st, content, {
    ...opts,
    size: opts.size ?? st.SIZE,
    line: opts.line ?? st.LINE,
    before: opts.before ?? 0,
    after: opts.after ?? 0,
  });
}

function defaultStyles(st: DocxStyleConfig): object {
  return {
    default: { document: { run: { font: st.FONT, size: st.SIZE } } },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        run: { font: st.FONT, size: st.SIZE, bold: true, color: st.HEADING_COLOR },
        paragraph: { alignment: AlignmentType.CENTER, spacing: { before: 240, after: 120 } },
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

function h1Impl(st: DocxStyleConfig, text: string, opts: ParagraphOpts = {}): Paragraph {
  return new Paragraph({
    ...paragraphOptions(st, { ...opts, align: AlignmentType.CENTER, noIndent: true, before: 240, after: 120, keepNext: true }, [
      runImpl(st, text.toUpperCase(), { ...opts, bold: true, color: st.HEADING_COLOR }),
    ]),
    heading: HeadingLevel.HEADING_1,
  });
}

function h2Impl(st: DocxStyleConfig, text: string, opts: ParagraphOpts = {}): Paragraph {
  return new Paragraph({
    ...paragraphOptions(st, { ...opts, before: 200, after: 80, keepNext: true }, [
      runImpl(st, text, { ...opts, bold: true, color: st.HEADING_COLOR }),
    ]),
    heading: HeadingLevel.HEADING_2,
  });
}

function h3Impl(st: DocxStyleConfig, text: string, opts: ParagraphOpts = {}): Paragraph {
  return new Paragraph({
    ...paragraphOptions(st, { ...opts, before: 140, after: 60, keepNext: true }, [
      runImpl(st, text, { ...opts, bold: true, italics: true, color: st.HEADING_COLOR }),
    ]),
    heading: HeadingLevel.HEADING_3,
  });
}

function blankImpl(st: DocxStyleConfig, opts: ParagraphOpts = {}): Paragraph {
  return new Paragraph({ spacing: { line: opts.line ?? st.LINE }, children: [runImpl(st, "", opts)] });
}

function pageBreakImpl(): Paragraph {
  return new Paragraph({ children: [new PageBreak()] });
}

function captionImpl(st: DocxStyleConfig, text: string, opts: ParagraphOpts = {}): Paragraph {
  return centeredImpl(st, text, {
    ...opts,
    before: opts.before ?? st.CAPTION_BEFORE,
    after: opts.after ?? st.CAPTION_AFTER,
    size: opts.size ?? st.SIZE,
  });
}

function tableCaptionImpl(st: DocxStyleConfig, text: string, opts: ParagraphOpts = {}): Paragraph {
  return paragraphImpl(st, text, {
    ...opts,
    align: AlignmentType.LEFT,
    noIndent: true,
    before: opts.before ?? st.TABLE_CAPTION_BEFORE,
    after: opts.after ?? st.TABLE_CAPTION_AFTER,
    size: opts.size ?? st.SIZE,
  });
}

function createCaptionCounterImpl(st: DocxStyleConfig, prefix: string, opts: CaptionCounterOpts = {}): CaptionCounter {
  let n = opts.start ?? 1;
  const fn = opts.table ? tableCaptionImpl : captionImpl;
  return {
    peek: () => n,
    caption: (text: string) => fn(st, `${prefix} ${n++} – ${text}`, opts),
  };
}

function placeholderImpl(st: DocxStyleConfig, label: string, opts: ParagraphOpts = {}): Paragraph {
  const border = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" };
  return new Paragraph({
    ...paragraphOptions(st, {
      ...opts,
      align: AlignmentType.CENTER,
      noIndent: true,
      before: opts.before ?? 120,
      after: opts.after ?? 60,
    }, [runImpl(st, label, { ...opts, italics: true, color: opts.color ?? "888888" })]),
    border: { top: border, bottom: border, left: border, right: border },
  });
}

function imageBlockImpl(st: DocxStyleConfig, imagePath: string | Buffer, width: number, height: number, opts: ParagraphOpts = {}): Paragraph {
  const data = Buffer.isBuffer(imagePath) ? imagePath : fs.readFileSync(imagePath);
  return new Paragraph({
    alignment: (opts.align as (typeof AlignmentType)[keyof typeof AlignmentType]) || AlignmentType.CENTER,
    spacing: {
      line: opts.line ?? st.LINE,
      before: opts.before ?? st.IMAGE_BEFORE,
      after: opts.after ?? st.IMAGE_AFTER,
    },
    keepNext: !!opts.keepNext,
    keepLines: opts.keepLines !== undefined ? opts.keepLines : true,
    children: [new ImageRun({ data, type: opts.imageType ?? "png", transformation: { width, height } } as ConstructorParameters<typeof ImageRun>[0])],
  });
}

function codeLineImpl(st: DocxStyleConfig, text: string, opts: ParagraphOpts = {}): Paragraph {
  return paragraphImpl(st, [{ text, font: opts.font ?? st.CODE_FONT, size: opts.size ?? st.SIZE_CODE }], {
    ...opts,
    noIndent: true,
    line: opts.line ?? st.LINE,
  });
}

export function toMathComponents(value: unknown): MathComponent[] {
  if (value === undefined || value === null) return [];
  if (Array.isArray(value)) return value.flatMap((item) => toMathComponents(item));
  if (typeof value === "string" || typeof value === "number") return [new MathRun(String(value))];
  return [value as MathComponent];
}

export function mathExpr(...parts: unknown[]): DocxMath {
  return new DocxMath({ children: parts.flatMap((part) => toMathComponents(part)) });
}

export function mathFraction(numerator: unknown, denominator: unknown): MathFraction {
  return new MathFraction({ numerator: toMathComponents(numerator), denominator: toMathComponents(denominator) });
}

export function mathSub(base: unknown, subScript: unknown): MathSubScript {
  return new MathSubScript({ children: toMathComponents(base), subScript: toMathComponents(subScript) });
}

export function mathSup(base: unknown, superScript: unknown): MathSuperScript {
  return new MathSuperScript({ children: toMathComponents(base), superScript: toMathComponents(superScript) });
}

export function mathSubSup(base: unknown, subScript: unknown, superScript: unknown): MathSubSuperScript {
  return new MathSubSuperScript({
    children: toMathComponents(base),
    subScript: toMathComponents(subScript),
    superScript: toMathComponents(superScript),
  });
}

export function mathParen(...parts: unknown[]): MathRoundBrackets {
  return new MathRoundBrackets({ children: parts.flatMap((part) => toMathComponents(part)) });
}

export function mathBracket(...parts: unknown[]): MathSquareBrackets {
  return new MathSquareBrackets({ children: parts.flatMap((part) => toMathComponents(part)) });
}

export function mathBrace(...parts: unknown[]): MathCurlyBrackets {
  return new MathCurlyBrackets({ children: parts.flatMap((part) => toMathComponents(part)) });
}

export function mathRoot(content: unknown, degree?: unknown): MathRadical {
  return new MathRadical({ children: toMathComponents(content), degree: degree === undefined ? undefined : toMathComponents(degree) });
}

export function mathSum(content: unknown, opts: { subScript?: unknown; superScript?: unknown } = {}): MathSum {
  return new MathSum({
    children: toMathComponents(content),
    subScript: opts.subScript ? toMathComponents(opts.subScript) : undefined,
    superScript: opts.superScript ? toMathComponents(opts.superScript) : undefined,
  });
}

export function mathIntegral(content: unknown, opts: { subScript?: unknown; superScript?: unknown } = {}): MathIntegral {
  return new MathIntegral({
    children: toMathComponents(content),
    subScript: opts.subScript ? toMathComponents(opts.subScript) : undefined,
    superScript: opts.superScript ? toMathComponents(opts.superScript) : undefined,
  });
}

function formulaMathImpl(st: DocxStyleConfig, content: unknown, opts: ParagraphOpts = {}): Paragraph {
  return new Paragraph(paragraphOptions(st, {
    ...opts,
    align: AlignmentType.CENTER,
    noIndent: true,
    before: opts.before ?? st.FORMULA_BEFORE,
    after: opts.after ?? st.FORMULA_AFTER,
  }, [mathExpr(content)]));
}

function formulaInlineImpl(st: DocxStyleConfig, label: string, mathContent: unknown, opts: ParagraphOpts = {}): Paragraph {
  return new Paragraph(paragraphOptions(st, opts, [runImpl(st, `${label} `, opts), mathExpr(mathContent)]));
}

function paragraphWithMathImpl(st: DocxStyleConfig, parts: Array<string | number | RunDescriptor | MathComponent>, opts: ParagraphOpts = {}): Paragraph {
  const children = parts.map((part): ParagraphChild => {
    if (typeof part === "string" || typeof part === "number") return runImpl(st, part, opts);
    if (part instanceof DocxMath) return part;
    if (Array.isArray(part) || isMathLike(part)) return mathExpr(part);
    const desc = part as RunDescriptor;
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
  return new Paragraph(paragraphOptions(st, opts, children));
}

function normalizeCell(cell: CellValue): CellDescriptor {
  if (typeof cell === "object" && cell !== null && !Array.isArray(cell) && !isMathLike(cell) && "text" in cell) {
    return cell as CellDescriptor;
  }
  return { text: cell };
}

function tableColumnWidths(st: DocxStyleConfig, count: number, widths?: number[]): number[] {
  if (Array.isArray(widths) && widths.length > 0) return widths;
  const base = Math.floor(st.CONTENT_WIDTH / count);
  const result = Array(count).fill(base);
  result[result.length - 1] += st.CONTENT_WIDTH - result.reduce((sum, width) => sum + width, 0);
  return result;
}

function tableCellImpl(st: DocxStyleConfig, cell: CellValue, width: number, opts: TableOpts & { bold?: boolean }): TableCell {
  const normalized = normalizeCell(cell);
  const content = normalized.text ?? "";
  const align = normalized.align ?? opts.align;
  const bold = normalized.bold !== undefined ? normalized.bold : opts.bold;
  const children = Array.isArray(content) || isMathLike(content)
    ? [paragraphWithMathImpl(st, Array.isArray(content) ? content as Array<string | number | RunDescriptor | MathComponent> : [content], {
      align,
      noIndent: true,
      line: opts.cellLine ?? st.TABLE_CELL_LINE,
      size: opts.fontSize,
    })]
    : [new Paragraph({
      alignment: align as (typeof AlignmentType)[keyof typeof AlignmentType],
      spacing: { line: opts.cellLine ?? st.TABLE_CELL_LINE, before: 0, after: 0 },
      children: [runImpl(st, String(content), {
        size: normalized.size ?? opts.fontSize,
        bold,
        italics: normalized.italics,
        color: normalized.color,
        font: normalized.font,
      })],
    })];
  return new TableCell({
    borders: opts.borders ?? st.BORDERS_ALL,
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    margins: st.TABLE_CELL_MARGINS,
    columnSpan: normalized.colspan,
    rowSpan: normalized.rowspan,
    shading: normalized.shading,
    children,
  });
}

function makeTableImpl(st: DocxStyleConfig, rows: CellValue[][], opts: TableOpts = {}): Table {
  const count = rows.reduce((max, row) => Math.max(max, row.length), 0);
  const effectiveStyle = opts.contentWidth ? { ...st, CONTENT_WIDTH: opts.contentWidth } : st;
  const widths = tableColumnWidths(effectiveStyle, count, opts.widths);
  const headerRow = opts.headerRow ?? 0;
  const defaultAlign = opts.align ?? AlignmentType.JUSTIFIED;
  const centerHeader = opts.centerHeader !== undefined ? opts.centerHeader : true;
  return new Table({
    width: { size: effectiveStyle.CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((row, rowIndex) => {
      const isHeader = rowIndex === headerRow;
      return new TableRow({
        children: row.map((cell, cellIndex) => tableCellImpl(st, cell, widths[cellIndex], {
          ...opts,
          align: isHeader && centerHeader ? AlignmentType.CENTER : defaultAlign,
          bold: isHeader,
          fontSize: opts.fontSize ?? st.SIZE,
          borders: opts.borders ?? st.BORDERS_ALL,
        })),
      });
    }),
  });
}

function makeFooterCityImpl(st: DocxStyleConfig, year: string | number): Footer {
  return new Footer({ children: [centeredImpl(st, `Таганрог ${year}`, { size: st.SIZE })] });
}

function makeFooterPageNumImpl(st: DocxStyleConfig): Footer {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: [PageNumber.CURRENT], font: st.FONT, size: st.SIZE })],
      }),
    ],
  });
}

function normalizeSubtitle(subtitle: string | string[] | undefined): string[] {
  if (Array.isArray(subtitle)) return subtitle;
  return subtitle === undefined || subtitle === null ? [] : [subtitle];
}

function makeTitlePageImpl(st: DocxStyleConfig, opts: TitlePageOpts): DocxSection {
  const children: Paragraph[] = [
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
    new Paragraph({ spacing: { line: st.LINE }, indent: { left: st.INDENT }, children: [runImpl(st, "Выполнил")] }),
    new Paragraph({
      spacing: { line: st.LINE },
      tabStops: st.TITLE_TAB_STOPS,
      children: [runImpl(st, `студент группы ${opts.group}`), new Tab(), runImpl(st, "_______________"), new Tab(), runImpl(st, opts.author)],
    }),
    blankImpl(st),
    new Paragraph({ spacing: { line: st.LINE }, indent: { left: st.INDENT }, children: [runImpl(st, "Принял")] }),
    new Paragraph({
      spacing: { line: st.LINE, after: 560 },
      tabStops: st.TITLE_TAB_STOPS,
      children: [runImpl(st, opts.teacherTitle ?? "Доцент кафедры"), new Tab(), runImpl(st, "_______________"), new Tab(), runImpl(st, opts.teacher)],
    }),
  ];
  return { properties: { page: st.PAGE }, footers: { default: makeFooterCityImpl(st, opts.year) }, children };
}

function makeContentSectionImpl(st: DocxStyleConfig, children: Array<Paragraph | Table>): DocxSection {
  return { properties: { page: st.PAGE }, footers: { default: makeFooterPageNumImpl(st) }, children };
}

function makeDocumentImpl(st: DocxStyleConfig, sections: DocxSection[], opts: { styles?: object } = {}): Document {
  return new Document({
    settings: { updateFields: true },
    styles: opts.styles ?? defaultStyles(st),
    sections,
  } as ConstructorParameters<typeof Document>[0]);
}

function saveDocumentImpl(doc: Document, outputPath: string): Promise<void> {
  return Packer.toBuffer(doc).then((buffer) => fs.promises.writeFile(outputPath, buffer));
}

export function createDocxGost(factoryOpts: DocxGostOptions = {}): DocxGostInstance {
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

export { TableOfContents };

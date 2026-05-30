import { LintIssue } from "./lint";
import { AlignmentType, BorderStyle, Document, Footer, Math as DocxMath, MathCurlyBrackets, MathFraction, MathIntegral, MathRadical, MathRoundBrackets, MathSquareBrackets, MathSubScript, MathSubSuperScript, MathSum, MathSuperScript, Paragraph, Table, TableOfContents, TabStopType, TextRun } from "docx";
type DocxBorder = {
    style: (typeof BorderStyle)[keyof typeof BorderStyle];
    size: number;
    color: string;
};
type DocxBorders = {
    top: DocxBorder;
    bottom: DocxBorder;
    left: DocxBorder;
    right: DocxBorder;
};
type TabStop = {
    type: (typeof TabStopType)[keyof typeof TabStopType];
    position: number;
};
type MathComponent = any;
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
    TABLE_CELL_MARGINS: {
        top: number;
        bottom: number;
        left: number;
        right: number;
    };
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
        size: {
            width: number;
            height: number;
        };
        margin: {
            top: number;
            bottom: number;
            left: number;
            right: number;
        };
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
    properties: {
        page: DocxStyleConfig["PAGE"];
    };
    footers: {
        default: Footer;
    };
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
    makeDocument(sections: DocxSection[], opts?: {
        styles?: object;
    }): Document;
    saveDocument(doc: Document, outputPath: string): Promise<void>;
    /**
     * Рендерит puml-строку → авторазмер → imageBlock + подпись.
     * Спредится прямо в children: `...await g.diagramBlock(puml, figures, "заголовок")`
     */
    diagramBlock(pumlSource: string, counter: CaptionCounter, captionText: string, opts?: DiagramBlockOpts): Promise<Paragraph[]>;
    /**
     * Анализирует children на пустые страницы и ГОСТ-нарушения.
     * Возвращает список проблем и выводит их в консоль.
     * Вызывать перед makeDocument.
     */
    lint(children: Array<Paragraph | Table>): LintIssue[];
}
export interface DiagramBlockOpts {
    /** DPI растеризации. По умолчанию: 150, для печати: 300 */
    dpi?: number;
    /** Максимальная ширина px. По умолчанию: 624 (ГОСТ A4) */
    maxWidth?: number;
    /** Переопределение skinparam */
    skinparams?: Record<string, string>;
}
export declare const DEFAULT_DOCX_STYLE: DocxStyleConfig;
export declare function createDocxStyle(opts?: DocxGostOptions["style"]): DocxStyleConfig;
export declare function toMathComponents(value: unknown): MathComponent[];
export declare function mathExpr(...parts: unknown[]): DocxMath;
export declare function mathFraction(numerator: unknown, denominator: unknown): MathFraction;
export declare function mathSub(base: unknown, subScript: unknown): MathSubScript;
export declare function mathSup(base: unknown, superScript: unknown): MathSuperScript;
export declare function mathSubSup(base: unknown, subScript: unknown, superScript: unknown): MathSubSuperScript;
export declare function mathParen(...parts: unknown[]): MathRoundBrackets;
export declare function mathBracket(...parts: unknown[]): MathSquareBrackets;
export declare function mathBrace(...parts: unknown[]): MathCurlyBrackets;
export declare function mathRoot(content: unknown, degree?: unknown): MathRadical;
export declare function mathSum(content: unknown, opts?: {
    subScript?: unknown;
    superScript?: unknown;
}): MathSum;
export declare function mathIntegral(content: unknown, opts?: {
    subScript?: unknown;
    superScript?: unknown;
}): MathIntegral;
export declare function createDocxGost(factoryOpts?: DocxGostOptions): DocxGostInstance;
export { TableOfContents };
export { lintChildren, printLintResults } from "./lint";
//# sourceMappingURL=docx-gost.d.ts.map
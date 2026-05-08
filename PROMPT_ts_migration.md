# Промпт: миграция lib на TypeScript + фабрики

## Контекст

В `/Users/valiev/Desktop/CODEX/lib/` лежат две JS-либы:
- `docx-gost.js` — генерация Word по ГОСТ 7.32
- `excel-gost.js` — генерация Excel-таблиц

Структура уже создана:
- `src/` — папка для TypeScript-исходников
- `tsconfig.json` — таргет ES2020, CommonJS, outDir=dist, rootDir=src
- `package.json` — main=dist/docx-gost.js, types=dist/docx-gost.d.ts

## Задача

Создать `src/docx-gost.ts` и `src/excel-gost.ts` — TypeScript-версии с фабричным паттерном. Старые `.js` файлы не удалять и не трогать.

---

## ЧАСТЬ 1: `src/docx-gost.ts`

### Принцип: фабрика закрывает стиль

```ts
const gost = createDocxGost();           // дефолтный ГОСТ-стиль
const gost = createDocxGost({ style: { font: "Arial" } });  // кастомный

gost.paragraph("Текст");
gost.h2("Заголовок");
gost.makeTable(rows, { widths: [3000, 6355] });
```

### Структура файла

```ts
import { ... } from "docx";
import fs from "fs";

// ── Типы ──────────────────────────────────────────────────
export interface DocxStyleConfig { ... }
export interface DocxGostOptions { style?: Partial<DocxStyleConfig> }
export interface DocxGostInstance { ... }

// ── Дефолты ───────────────────────────────────────────────
const DEFAULT_STYLE: DocxStyleConfig = { ... };

// ── Фабрика стиля ─────────────────────────────────────────
export function createDocxStyle(opts?: Partial<DocxStyleConfig>): DocxStyleConfig { ... }

// ── Math-хелперы (pure, без стиля) — экспортируются отдельно
export function toMathComponents(...) { ... }
export function mathFraction(...) { ... }
// ... все math*

// ── Фабрика либы ──────────────────────────────────────────
export function createDocxGost(opts?: DocxGostOptions): DocxGostInstance { ... }
```

### Интерфейс `DocxStyleConfig`

Все поля из `DEFAULT_DOCX_STYLE` в `docx-gost.js`:

```ts
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
  BORDER_SINGLE: { style: BorderStyle; size: number; color: string };
  BORDERS_ALL: { top: object; bottom: object; left: object; right: object };
  TITLE_TAB_STOPS: { type: TabStopType; position: number }[];
}
```

### Типы опций для методов

```ts
type Content = string | number | Array<string | number | RunDescriptor>;

interface RunDescriptor {
  text: string;
  size?: number;
  bold?: boolean;
  italics?: boolean;
  allCaps?: boolean;
  underline?: boolean | object;
  color?: string;
  font?: string;
}

interface ParagraphOpts {
  align?: string;
  noIndent?: boolean;
  indent?: number;
  line?: number;
  before?: number;
  after?: number;
  bold?: boolean;
  italics?: boolean;
  size?: number;
  color?: string;
  font?: string;
  keepNext?: boolean;
  keepLines?: boolean;
}

interface TableOpts {
  widths?: number[];
  headerRow?: number;
  align?: string;
  centerHeader?: boolean;
  fontSize?: number;
  cellLine?: number;
  contentWidth?: number;
  borders?: object;
}

interface CaptionCounterOpts {
  start?: number;
  table?: boolean;
}

interface CaptionCounter {
  peek: () => number;
  caption: (text: string) => Paragraph;
}

type CellValue = string | number | CellDescriptor | MathComponent;
interface CellDescriptor {
  text: CellValue | CellValue[];
  align?: string;
  bold?: boolean;
  italics?: boolean;
  size?: number;
  color?: string;
  font?: string;
  colspan?: number;
  rowspan?: number;
  shading?: object;
}

interface TitlePageOpts {
  workType: string;
  subtitle?: string | string[];
  discipline: string;
  group: string;
  author: string;
  teacher: string;
  teacherTitle?: string;
  year: number | string;
}

type MathComponent = object; // docx math types
```

### Интерфейс `DocxGostInstance`

```ts
export interface DocxGostInstance {
  readonly style: DocxStyleConfig;

  // Текст
  run(text: string | number, opts?: ParagraphOpts): TextRun;
  paragraph(content: Content, opts?: ParagraphOpts): Paragraph;
  centered(content: Content, opts?: ParagraphOpts): Paragraph;
  paragraphWithMath(parts: Array<string | number | RunDescriptor | MathComponent>, opts?: ParagraphOpts): Paragraph;

  // Заголовки
  h1(text: string, opts?: ParagraphOpts): Paragraph;
  h2(text: string, opts?: ParagraphOpts): Paragraph;
  h3(text: string, opts?: ParagraphOpts): Paragraph;

  // Структурные элементы
  blank(opts?: ParagraphOpts): Paragraph;
  pageBreak(): Paragraph;
  caption(text: string, opts?: ParagraphOpts): Paragraph;
  tableCaption(text: string, opts?: ParagraphOpts): Paragraph;
  createCaptionCounter(prefix: string, opts?: CaptionCounterOpts): CaptionCounter;
  placeholder(label: string, opts?: ParagraphOpts): Paragraph;
  imageBlock(imagePath: string | Buffer, width: number, height: number, opts?: ParagraphOpts): Paragraph;

  // Код
  codeLine(text: string, opts?: ParagraphOpts): Paragraph;
  codeBlock(lines: string[], opts?: ParagraphOpts): Paragraph[];

  // Формулы (нужны стиль для отступов)
  formulaMath(content: MathComponent | MathComponent[], opts?: ParagraphOpts): Paragraph;
  formulaInline(label: string, mathContent: MathComponent | MathComponent[], opts?: ParagraphOpts): Paragraph;

  // Таблицы
  makeTable(rows: CellValue[][], opts?: TableOpts): Table;

  // Документ
  makeTitlePage(opts: TitlePageOpts): object;
  makeContentSection(children: Array<Paragraph | Table>, opts?: object): object;
  makeDocument(sections: object[], opts?: object): Document;
  saveDocument(doc: Document, outputPath: string): Promise<void>;
}
```

### Реализация фабрики

```ts
export function createDocxGost(factoryOpts: DocxGostOptions = {}): DocxGostInstance {
  const st = createDocxStyle(factoryOpts.style);

  // Биндим стиль: при каждом вызове метода стиль из фабрики является дефолтом,
  // но пользователь может переопределить его через opts.style в конкретном вызове
  const s = (opts: ParagraphOpts = {}): ParagraphOpts & { _st: DocxStyleConfig } =>
    ({ ...opts, _st: (opts as any)._st ?? st });

  return {
    style: st,
    run:              (text, opts = {}) => runImpl(st, text, opts),
    paragraph:        (content, opts = {}) => paragraphImpl(st, content, opts),
    centered:         (content, opts = {}) => centeredImpl(st, content, opts),
    paragraphWithMath:(parts, opts = {}) => paragraphWithMathImpl(st, parts, opts),
    h1:               (text, opts = {}) => h1Impl(st, text, opts),
    h2:               (text, opts = {}) => h2Impl(st, text, opts),
    h3:               (text, opts = {}) => h3Impl(st, text, opts),
    blank:            (opts = {}) => blankImpl(st, opts),
    pageBreak:        () => pageBreakImpl(),
    caption:          (text, opts = {}) => captionImpl(st, text, opts),
    tableCaption:     (text, opts = {}) => tableCaptionImpl(st, text, opts),
    createCaptionCounter: (prefix, opts = {}) => createCaptionCounterImpl(st, prefix, opts),
    placeholder:      (label, opts = {}) => placeholderImpl(st, label, opts),
    imageBlock:       (src, w, h, opts = {}) => imageBlockImpl(st, src, w, h, opts),
    codeLine:         (text, opts = {}) => codeLineImpl(st, text, opts),
    codeBlock:        (lines, opts = {}) => lines.map(l => codeLineImpl(st, l, opts)),
    formulaMath:      (content, opts = {}) => formulaMathImpl(st, content, opts),
    formulaInline:    (label, math, opts = {}) => formulaInlineImpl(st, label, math, opts),
    makeTable:        (rows, opts = {}) => makeTableImpl(st, rows, opts),
    makeTitlePage:    (opts) => makeTitlePageImpl(st, opts),
    makeContentSection: (children, opts = {}) => makeContentSectionImpl(st, children, opts),
    makeDocument:     (sections, opts = {}) => makeDocumentImpl(st, sections, opts),
    saveDocument:     (doc, path) => saveDocumentImpl(doc, path),
  };
}
```

### Внутренние функции (`*Impl`)

Это приватные функции модуля (не экспортируются), принимают `st: DocxStyleConfig` первым параметром. Логика полностью совпадает с `docx-gost.js`, только:
- Убрать `styleFromOpts` — стиль передаётся явно как `st`
- Убрать `opts.style` — стиль уже в `st`
- Все `const style = styleFromOpts(opts)` → просто используй `st`

Пример:
```ts
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
```

### Math-хелперы (экспортируемые standalone)

Не нужен стиль, экспортировать напрямую без изменений:

```ts
export function toMathComponents(value: unknown): object[] { ... }
export function mathExpr(...parts: unknown[]): Math { ... }
export function mathFraction(numerator: unknown, denominator: unknown): MathFraction { ... }
export function mathSub(base: unknown, sub: unknown): MathSubScript { ... }
export function mathSup(base: unknown, sup: unknown): MathSuperScript { ... }
export function mathSubSup(base: unknown, sub: unknown, sup: unknown): MathSubSuperScript { ... }
export function mathParen(...parts: unknown[]): MathRoundBrackets { ... }
export function mathBracket(...parts: unknown[]): MathSquareBrackets { ... }
export function mathBrace(...parts: unknown[]): MathCurlyBrackets { ... }
export function mathRoot(content: unknown, degree?: unknown): MathRadical { ... }
export function mathSum(content: unknown, opts?: { subScript?: unknown; superScript?: unknown }): MathSum { ... }
export function mathIntegral(content: unknown, opts?: { subScript?: unknown; superScript?: unknown }): MathIntegral { ... }
```

### Что экспортировать в конце

```ts
export {
  // Фабрики
  createDocxStyle,
  createDocxGost,

  // Math standalone
  toMathComponents, mathExpr,
  mathFraction, mathSub, mathSup, mathSubSup,
  mathParen, mathBracket, mathBrace, mathRoot, mathSum, mathIntegral,

  // Константы (для совместимости)
  DEFAULT_STYLE as DEFAULT_DOCX_STYLE,
};

// Re-export TableOfContents из docx (нужен в build_report)
export { TableOfContents } from "docx";
```

---

## ЧАСТЬ 2: `src/excel-gost.ts`

Аналогично. Фабрика:

```ts
const xl = createExcelGost();                          // дефолт (жёлтая палитра v2)
const xl = createExcelGost({ headerFill: "D9EAF7" }); // кастом (синяя палитра v4)

xl.setValue(sheet, "A1", "Заголовок", xl.HEADER_STYLE);
xl.makeTable(sheet, rows, { startRow: 3 });
```

### Интерфейс `ExcelStyleConfig`

```ts
export interface ExcelStyleConfig {
  HEADER_FILL: string;
  RESULT_FILL: string;
  ALT_ROW_FILL: string;
  REGRET_FILL: string;
  FONT_NAME: string;
  FONT_SIZE: number;
  TAB_COLOR: string;
}
```

### Интерфейс `ExcelGostInstance`

```ts
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

  // Pure formula helpers (also exported standalone)
  productFormula(addresses: string[], exponent: number): string;
  absoluteRef(address: string): string;
  weightedSumFormula(valueAddresses: string[], weightAddresses: string[]): string;
  listFormulaRange(sheetName: string, addresses: string[]): string;
}

interface AhpBlockResult {
  weightColumn: string;
  weightStartRow: number;
  weightEndRow: number;
  osCell: string;
  nextRow: number;
}
```

### Фабрика

```ts
export function createExcelStyle(opts: Partial<ExcelStyleConfig> = {}): ExcelStyleConfig {
  return {
    HEADER_FILL:  opts.HEADER_FILL  ?? "FFF2CC",
    RESULT_FILL:  opts.RESULT_FILL  ?? "D9EAD3",
    ALT_ROW_FILL: opts.ALT_ROW_FILL ?? "FFF9F0",
    REGRET_FILL:  opts.REGRET_FILL  ?? "FCE4D6",
    FONT_NAME:    opts.FONT_NAME    ?? "Times New Roman",
    FONT_SIZE:    opts.FONT_SIZE    ?? 12,
    TAB_COLOR:    opts.TAB_COLOR    ?? "FFF2CC",
  };
}

export function createExcelGost(opts: Partial<ExcelStyleConfig> = {}): ExcelGostInstance {
  const st = createExcelStyle(opts);
  const HEADER_STYLE: CellStyleOptions = { bold: true, border: true, fill: st.HEADER_FILL, align: "center" };
  const RESULT_STYLE: CellStyleOptions = { bold: true, border: "medium", fill: st.RESULT_FILL, align: "center" };

  return {
    style: st,
    HEADER_STYLE,
    RESULT_STYLE,
    styleCell:          (cell, options = {}) => styleCellImpl(cell, options, st),
    setValue:           (sheet, addr, val, options = {}) => setValueImpl(sheet, addr, val, options, st),
    setFormula:         (sheet, addr, formula, options = {}) => setFormulaImpl(sheet, addr, formula, options, st),
    applyBaseSheetSetup:(sheet, widths) => applyBaseSheetSetupImpl(sheet, widths),
    mergeTitle:         (sheet, range, text) => mergeTitleImpl(sheet, range, text, st),
    setFreeze:          (sheet, ySplit, activeCell) => setFreezeImpl(sheet, ySplit, activeCell),
    applyZebraStriping: (sheet, sr, er, sc, ec, hr) => applyZebraStripingImpl(sheet, sr, er, sc, ec, hr, st),
    getCellText:        (cell) => getCellTextImpl(cell),
    autoFitSheet:       (sheet) => autoFitSheetImpl(sheet),
    formatTableRange:   (sheet, sr, er, sc, ec, hr) => formatTableRangeImpl(sheet, sr, er, sc, ec, hr, st),
    addAhpMatrixBlock:  (sheet, sr, title, labels, matrix, ri) => addAhpMatrixBlockImpl(sheet, sr, title, labels, matrix, ri, st),
    addContentsSheet:   (wb, names) => addContentsSheetImpl(wb, names, st),
    productFormula,
    absoluteRef,
    weightedSumFormula,
    listFormulaRange,
  };
}
```

Внутренние `*Impl` функции — логика из `excel-gost.js`, только первым параметром берут `st: ExcelStyleConfig` вместо обращения к глобальным константам.

### Что экспортировать

```ts
export {
  createExcelStyle,
  createExcelGost,
  // Pure helpers standalone
  productFormula,
  absoluteRef,
  weightedSumFormula,
  listFormulaRange,
};
```

---

## Установка зависимостей

Перед написанием файлов:

```bash
cd /Users/valiev/Desktop/CODEX/lib
# TypeScript уже должен быть в devDependencies
# Если node_modules — симлинк, установить локально:
npm install --save-dev typescript
npm install  # установит docx и exceljs если нет
```

## Сборка и проверка

```bash
cd /Users/valiev/Desktop/CODEX/lib
npx tsc --noEmit   # проверить типы без вывода файлов
npx tsc            # собрать в dist/
```

Должно появиться:
- `dist/docx-gost.js`
- `dist/docx-gost.d.ts`
- `dist/excel-gost.js`
- `dist/excel-gost.d.ts`

## Правила

1. Не удалять и не изменять `docx-gost.js` и `excel-gost.js` — они остаются как есть.
2. Все `*Impl` функции — приватные (не экспортируются).
3. `strict: true` — никаких `any` кроме мест где это неизбежно (legacy docx types).
4. Math-хелперы — экспортируются standalone, не через фабрику (они pure).
5. После сборки сделать коммит: `git add src/ dist/ && git commit -m "feat: migrate to TypeScript with factory pattern"`

# @magomed-cmd/gost

Библиотека для программной генерации учебных отчётов по ГОСТ 7.32.

Форматы:
- `.docx` через `docx`
- `.xlsx` через `exceljs`
- диаграммы через `node-plantuml`

Главная идея: фабрики `createDocxGost()`, `createExcelGost()`, `createPlantUmlGost()` создают независимые экземпляры со своими настройками стиля. Нет глобального состояния — один процесс может собирать несколько документов с разными параметрами.

## Быстрый выбор

| Задача | Использовать |
| --- | --- |
| Собрать Word-отчёт | `createDocxGost()` из `@magomed-cmd/gost/docx` |
| Вставить рисунок с автоподписью | `g.diagramBlock()` или `imageBlock()` + `figures.caption()` |
| Проверить документ на пустые страницы | `g.lint(children)` перед `makeDocument()` |
| Нарисовать UML-диаграмму в коде | `createPlantUmlGost()` из `@magomed-cmd/gost/plantuml` |
| Авторазмер PNG под ширину ГОСТ | `autoImageSize(buffer)` |
| Сгенерировать Excel-файл | `createExcelGost()` из `@magomed-cmd/gost/excel` |
| Построить блок МАИ в Excel | `addAhpMatrixBlock()` |

## Установка и сборка

```bash
cd gost
npm install
npm run build
npm test
```

После сборки появляются CommonJS-файлы в `dist/`. Subpath entrypoints пакета:

```js
const { createDocxGost }     = require("@magomed-cmd/gost/docx");
const { createExcelGost }    = require("@magomed-cmd/gost/excel");
const { createPlantUmlGost } = require("@magomed-cmd/gost/plantuml");
```

Или через относительный путь (если пакет не установлен глобально):

```js
const { createDocxGost }     = require("../gost/dist/docx-gost");
const { createExcelGost }    = require("../gost/dist/excel-gost");
const { createPlantUmlGost } = require("../gost/dist/plantuml-gost");
```

---

## DOCX

### Минимальный пример

```js
const path = require("path");
const fs   = require("fs");
const { createDocxGost, TableOfContents } = require("@magomed-cmd/gost/docx");

const g = createDocxGost();
const {
  h1, h2, paragraph, makeTable, imageBlock,
  createCaptionCounter, pageBreak,
  makeTitlePage, makeContentSection, makeDocument, saveDocument,
} = g;

const figures = createCaptionCounter("Рисунок");
const tables  = createCaptionCounter("Таблица", { table: true });

const children = [
  h1("Содержание"),
  new TableOfContents("Содержание", { hyperlink: false, headingStyleRange: "1-3" }),
  pageBreak(),

  h1("1 Практическая часть"),
  h2("1.1 Исходные данные"),
  paragraph(`Данные приведены в таблице ${tables.peek()}.`),
  tables.caption("исходная матрица"),
  makeTable([
    ["", "b1", "b2"],
    ["a1", 3, 5],
    ["a2", 4, 2],
  ]),

  paragraph(`График показан на рисунке ${figures.peek()}.`),
  imageBlock(fs.readFileSync("graph.png"), 600, 420),
  figures.caption("графический метод"),
];

// Проверка на пустые страницы и нарушения ГОСТ перед сборкой
g.lint(children);

const doc = makeDocument([
  makeTitlePage({
    workType: "ОТЧЁТ",
    subtitle: ["о выполнении лабораторной работы №1", "на тему: «Пример»"],
    discipline: "по дисциплине: «Архитектура ПО»",
    group: "КТбо3-7",
    author: "Валиев М. М.",
    teacher: "Лутай В. Н.",
    teacherTitle: "Доцент кафедры",
    year: 2026,
  }),
  makeContentSection(children),
]);

await saveDocument(doc, path.join(__dirname, "report.docx"));
```

---

### Настройка стиля

`createDocxGost(opts)` принимает опции стиля. Все поля необязательны — базовые значения соответствуют ГОСТ 7.32:

| Alias | Поле | Значение | Смысл |
| --- | --- | --- | --- |
| `font` | `FONT` | Times New Roman | основной шрифт |
| `size` | `SIZE` | 28 | основной размер (14 pt) |
| `sizeCode` | `SIZE_CODE` | 24 | шрифт кода (12 pt) |
| `indent` | `INDENT` | 709 | красная строка (~1.25 см) |
| `line` | `LINE` | 360 | межстрочный интервал |
| `contentWidth` | `CONTENT_WIDTH` | 9355 | ширина контента в DXA |
| `codeFont` | `CODE_FONT` | Courier New | моноширинный шрифт |

```js
const g = createDocxGost({ style: { sizeCode: 20 } });
```

Размеры шрифтов в `docx` задаются в half-points: `28` = 14 pt, `24` = 12 pt.

---

### Текстовые элементы

#### `paragraph(content, opts)`

Обычный абзац. По умолчанию: выравнивание по ширине, красная строка, `keepLines: true`.

```js
paragraph("Текст абзаца.");
paragraph("Без красной строки.", { noIndent: true });
paragraph([
  "Метод ",
  { text: "Джонсона", bold: true },
  " применяется для упорядочивания работ.",
]);
```

Опции: `align`, `noIndent`, `bold`, `italics`, `size`, `color`, `font`, `keepNext`, `before`, `after`.

#### `centered(content, opts)`

То же, что `paragraph()`, но по центру без красной строки.

#### `h1(text)` / `h2(text)` / `h3(text)`

Заголовки. `h1` автоматически переводит текст в uppercase и центрирует. Все заголовки попадают в `TableOfContents`.

#### `blank()` / `pageBreak()`

Пустой абзац и принудительный разрыв страницы.

---

### Подписи и счётчики

#### `createCaptionCounter(prefix, opts)`

```js
const figures = createCaptionCounter("Рисунок");
const tables  = createCaptionCounter("Таблица", { table: true });
```

| Метод | Что делает |
| --- | --- |
| `peek()` | возвращает текущий номер без инкремента |
| `caption(text)` | создаёт подпись и инкрементирует счётчик |

Паттерн для рисунка:
```js
paragraph(`На рисунке ${figures.peek()} показан результат.`);
imageBlock(buffer, 600, 420);
figures.caption("результат запуска программы");
```

Паттерн для таблицы (подпись — над таблицей):
```js
paragraph(`Данные приведены в таблице ${tables.peek()}.`);
tables.caption("исходные данные");
makeTable(rows);
```

#### `placeholder(label)`

Серый блок-заглушка когда изображение ещё не готово.

```js
placeholder("[Рисунок — скриншот главного окна]");
```

---

### Рисунки

#### `imageBlock(pathOrBuffer, width, height, opts)`

```js
imageBlock(fs.readFileSync("graph.png"), 600, 420);
imageBlock(fs.readFileSync("photo.jpg"), 400, 300, { imageType: "jpg" });
```

По умолчанию `keepNext: true` — изображение держится со следующим абзацем (подписью), что предотвращает разрыв страницы между рисунком и подписью.

#### `diagramBlock(pumlSource, counter, captionText, opts?)`

Комбо-метод: рендерит PlantUML строку → авторазмер → `imageBlock` + подпись. Спредится прямо в `children`.

```js
const { createPlantUmlGost } = require("@magomed-cmd/gost/plantuml");
const p = createPlantUmlGost();

const children = [
  paragraph(`На рисунке ${figures.peek()} представлена диаграмма классов.`),
  ...await g.diagramBlock(
    p.classDiagram()
      .interface("RepairStrategy", ["+ getType(): String", "+ repair(): void"])
      .class("ElectricianRepair")
      .implements("ElectricianRepair", "RepairStrategy")
      .build(),
    figures,
    "диаграмма классов системы"
  ),
];
```

Опции `DiagramBlockOpts`:

| Поле | По умолчанию | Смысл |
| --- | --- | --- |
| `dpi` | 150 | качество (300 для печати) |
| `maxWidth` | 624 | максимальная ширина px (ГОСТ A4) |
| `skinparams` | — | переопределение skinparam PlantUML |

---

### Листинги кода

```js
codeLine("npm run build");
...codeBlock(fs.readFileSync("app.js", "utf8").split("\n"));
```

---

### Таблицы

```js
makeTable([
  ["Критерий", "Значение"],
  ["MM", "E1"],
  [{ text: "Итого", bold: true }, 42],
], { widths: [4000, 5355] });
```

Ячейка может быть строкой, числом, объектом с полями `text`, `bold`, `align`, `colspan`, `rowspan`, `shading`.

---

### Математика

Функции импортируются отдельно — они не зависят от экземпляра:

```js
const { createDocxGost, mathFraction, mathSub, mathSup } = require("@magomed-cmd/gost/docx");
```

```js
paragraphWithMath(["Стратегия ", mathSub("a", "1"), " оптимальна."]);
formulaMath(["ν = ", mathFraction("48", "13"), " ≈ 3.692"]);
formulaMath([mathSum(mathSub("x", "i"), { subScript: "i=1", superScript: "n" })]);
```

Доступны: `mathFraction`, `mathSub`, `mathSup`, `mathSubSup`, `mathParen`, `mathBracket`, `mathBrace`, `mathRoot`, `mathSum`, `mathIntegral`, `mathExpr`.

---

### Lint: проверка документа

`g.lint(children)` анализирует массив элементов **до** сборки документа и выводит предупреждения в консоль.

```js
g.lint(children);
const doc = makeDocument([makeTitlePage(opts), makeContentSection(children)]);
```

Пример вывода:
```
  ⚠ [WARN] pos  14  Два pageBreak подряд на позиции 14–15 — пустая страница.
  ⚠ [WARN] pos  47  image на позиции 47 отделён от подписи явным pageBreak — рисунок и подпись на разных страницах.
  ⚠ [WARN] pos  83  image на позиции 83 не имеет caption следом (ГОСТ требует подпись).
  ℹ [INFO] pos  12  blank() на позиции 12 перед pageBreak — лишний, можно убрать.

lintChildren: 3 предупреждений, 1 замечание.
```

Полный список правил:

| Код | Уровень | Условие |
| --- | --- | --- |
| `empty-page-start` | warn | `pageBreak` первым элементом |
| `empty-page-end` | warn | `pageBreak` последним элементом |
| `consecutive-page-breaks` | warn | два `pageBreak` подряд |
| `page-break-blank-page-break` | warn | `pageBreak → blank(и) → pageBreak` |
| `heading-before-page-break` | warn | `h1/h2/h3` прямо перед `pageBreak` |
| `image-caption-split` | warn | `image → pageBreak → caption` |
| `table-caption-split` | warn | `tableCaption → pageBreak → table` |
| `image-without-caption` | warn | `image` без `caption` следом |
| `placeholder-left` | warn | `placeholder` в финальном документе |
| `blank-before-page-break` | info | `blank` прямо перед `pageBreak` |
| `too-many-blanks` | info | 3+ `blank` подряд |

Для программного использования:

```js
const { lintChildren, printLintResults } = require("@magomed-cmd/gost/docx");

const issues = lintChildren(children);
const warnings = issues.filter(i => i.level === "warn");
if (warnings.length > 0) process.exit(1);
```

---

### Сборка документа

```js
const doc = makeDocument([
  makeTitlePage({
    workType: "ОТЧЁТ",
    subtitle: ["о выполнении лабораторной работы №4", "на тему: «TramLine»"],
    discipline: "по дисциплине: «Архитектура и технологии разработки ПО»",
    group: "КТбо3-7",
    author: "Валиев М. М.",
    teacher: "Лутай В. Н.",
    year: 2026,
  }),
  makeContentSection(children),
]);

await saveDocument(doc, path.join(__dirname, "report.docx"));
```

---

## PlantUML

### Минимальный пример

```js
const { createPlantUmlGost } = require("@magomed-cmd/gost/plantuml");

const p = createPlantUmlGost(); // dpi: 150 по умолчанию

const png = await p.render(
  p.classDiagram()
    .interface("TramObserver", ["+ onTramEvent(e: TramEvent): void"])
    .class("MainFrame", ["- dispatcher: Dispatcher"])
    .implements("MainFrame", "TramObserver")
    .build()
);

// png — Buffer, вставляется через imageBlock
imageBlock(png, 500, 400);
figures.caption("диаграмма классов системы TramLine");
```

### `createPlantUmlGost(opts?)`

| Поле | По умолчанию | Смысл |
| --- | --- | --- |
| `dpi` | 150 | разрешение PNG |

Требует установленной Java и пакета `node-plantuml` (входит в зависимости).

### Builders

Все builders поддерживают `.add(line)` как escape hatch для любой строки PlantUML, которую builder не покрывает.

#### `p.useCaseDiagram()`

```js
p.useCaseDiagram()
  .actor("Водитель", "D")
  .actor("Диспетчер", "Disp")
  .usecase("Сообщить о поломке", "UC1")
  .usecase("Выслать бригаду", "UC2")
  .arrow("D", "UC1")
  .arrow("Disp", "UC2")
  .include("UC1", "UC2")
  .package("Система", b => b.usecase("Внутреннее действие"))
  .build()
```

Методы: `actor`, `usecase`, `arrow`, `extend`, `include`, `package`, `blank`, `add`.

#### `p.classDiagram()`

```js
p.classDiagram()
  .interface("RepairStrategy", ["+ getType(): String", "+ getRepairTimeMs(): int"])
  .class("ElectricianRepair", ["+ getType(): String"])
  .class("MechanicRepair",    ["+ getType(): String"])
  .implements("ElectricianRepair", "RepairStrategy")
  .implements("MechanicRepair", "RepairStrategy")
  .enum("TramState", ["ON_LINE", "IN_DEPOT", "BROKEN", "RESERVE"])
  .association("Dispatcher", "TramRoute", "управляет")
  .composition("TramRoute", "Tram")
  .dependency("Dispatcher", "RepairStrategy", "использует")
  .package("tramline.model", b => b.class("Tram").class("TramRoute"))
  .build()
```

`build()` автоматически добавляет `hide circle`.

Методы: `class`, `abstract`, `interface`, `enum`, `extends`, `implements`, `association`, `dependency`, `composition`, `aggregation`, `package`, `blank`, `add`.

#### `p.sequenceDiagram()`

```js
p.sequenceDiagram()
  .actor("Пользователь", "U")
  .participant("MainFrame", "UI")
  .participant("Dispatcher", "D")
  .divider("Поломка трамвая")
  .arrow("U", "UI", "нажать «Поломка»")
  .arrow("UI", "D", "triggerBreakdown()")
  .activate("D")
  .arrow("D", "D", "doBreakdown()")
  .return("D", "UI", "publish(BREAKDOWN)")
  .deactivate("D")
  .note("Всё в dispatcher-thread", "right", "D")
  .build()
```

Методы: `participant`, `actor`, `arrow`, `return`, `activate`, `deactivate`, `note`, `divider`, `group`, `blank`, `add`.

#### `p.stateDiagram()`

```js
p.stateDiagram()
  .initial("ON_LINE")
  .state("ON_LINE", "На линии")
  .state("BROKEN",  "Сломан")
  .state("IN_DEPOT","В парке")
  .state("RESERVE", "Резерв")
  .transition("ON_LINE", "BROKEN",  "поломка",    "down")
  .transition("BROKEN",  "IN_DEPOT","отремонтирован")
  .transition("IN_DEPOT","ON_LINE", "выход на линию")
  .transition("ON_LINE", "RESERVE", "избыток")
  .transition("RESERVE", "ON_LINE", "нехватка")
  .final("IN_DEPOT")
  .build()
```

Методы: `state`, `transition`, `initial`, `final`, `blank`, `add`.

`transition` принимает опциональное направление: `"down"`, `"up"`, `"left"`, `"right"` — помогает избежать пересечений стрелок.

### `autoImageSize(buffer, maxWidth?)`

Читает реальные размеры PNG из буфера и масштабирует пропорционально.

```js
const { autoImageSize } = require("@magomed-cmd/gost/plantuml");

const png = await p.render(puml);
const { width, height } = autoImageSize(png);       // maxWidth = 624 (ГОСТ A4)
const { width, height } = autoImageSize(png, 400);  // своя ширина
imageBlock(png, width, height);
```

Максимальная ширина контента ГОСТ A4 = **624px** (9355 DXA при 150 DPI).

### `wrapWithSkin(source, dpi?, opts?)`

Автоматически добавляет базовые skinparam к любой puml-строке. Вызывается внутри `render()`, но доступен публично.

```js
const { wrapWithSkin } = require("@magomed-cmd/gost/plantuml");

const styled = wrapWithSkin(rawPuml, 300, {
  skinparams: { backgroundColor: "white" }
});
```

Базовые skinparam: Times New Roman 12pt, dpi 150, Shadowing false, nodesep 60, ranksep 50.

---

## Excel

### Минимальный пример

```js
const ExcelJS = require("exceljs");
const { createExcelGost } = require("@magomed-cmd/gost/excel");

const workbook = new ExcelJS.Workbook();
const x = createExcelGost({ fontSize: 14, headerFill: "FFF2CC" });

const sheet = workbook.addWorksheet("Расчёт", {
  properties: { tabColor: { argb: x.style.TAB_COLOR } },
});

x.applyBaseSheetSetup(sheet, [24, 18, 18, 18]);
x.mergeTitle(sheet, "A1:D1", "ИСХОДНАЯ МАТРИЦА");

x.setValue(sheet, "A3", "Стратегия", x.HEADER_STYLE);
x.setValue(sheet, "B3", "b1", x.HEADER_STYLE);
x.setValue(sheet, "B4", 3, { border: true });
x.setFormula(sheet, "B10", "SUM(B4:B9)", { border: true, numFmt: "0.000" });

x.formatTableRange(sheet, 3, 10, 1, 2, 3);
x.autoFitSheet(sheet);

await workbook.xlsx.writeFile("report.xlsx");
```

### Настройка стиля

```js
const x = createExcelGost({
  fontSize: 14,          // alias для FONT_SIZE
  fontName: "Arial",     // alias для FONT_NAME
  headerFill: "FFF2CC",  // alias для HEADER_FILL
  resultFill: "D9EAD3",  // alias для RESULT_FILL
  tabColor: "FFF2CC",    // alias для TAB_COLOR
});
```

### Методы

| Метод | Смысл |
| --- | --- |
| `applyBaseSheetSetup(sheet, widths)` | базовая разметка листа: высоты, ориентация, ширины колонок |
| `mergeTitle(sheet, range, text)` | объединить диапазон и написать крупный заголовок |
| `setValue(sheet, addr, value, style?)` | записать значение и применить стиль |
| `setFormula(sheet, addr, formula, style?)` | записать формулу без `=` |
| `styleCell(cell, opts)` | применить стиль к ячейке |
| `setFreeze(sheet, ySplit, activeCell)` | заморозить строки сверху |
| `formatTableRange(sheet, r1, r2, c1, c2, hRow)` | оформить диапазон как таблицу |
| `applyZebraStriping(sheet, r1, r2, c1, c2, headerRows)` | зебра-заливка чётных строк |
| `autoFitSheet(sheet)` | автоширина колонок и высота строк — вызывать в конце листа |
| `addAhpMatrixBlock(sheet, row, title, labels, matrix, ri)` | блок МАИ с весами и ОС |
| `addContentsSheet(workbook, names)` | лист «Содержание» со ссылками |

### Блок МАИ

```js
const result = x.addAhpMatrixBlock(
  sheet, 5,
  "Матрица парных сравнений",
  ["К1", "К2", "К3"],
  [[1, 3, 5], [1/3, 1, 2], [1/5, 1/2, 1]],
  0.58  // случайный индекс
);
// result: { weightColumn, weightStartRow, weightEndRow, osCell, nextRow }
```

---

## Тесты

```bash
npm test               # запустить все тесты
npm run test:coverage  # с отчётом покрытия
```

80 тестов в 5 файлах:

| Файл | Что проверяется |
| --- | --- |
| `plantuml-size.test.ts` | `autoImageSize` — масштабирование, пропорции |
| `plantuml-render.test.ts` | `wrapWithSkin` — skinparam injection и переопределение |
| `plantuml-builders.test.ts` | все 4 builder-а и их методы |
| `docx-gost.test.ts` | публичный API, `CaptionCounter`, `makeDocument` |
| `lint.test.ts` | все 11 lint-правил |

---

## Полный список экспортов

### `@magomed-cmd/gost/docx`

```js
// Фабрика
createDocxGost(opts?)

// Математика (не зависят от экземпляра)
mathExpr, mathFraction, mathSub, mathSup, mathSubSup,
mathParen, mathBracket, mathBrace, mathRoot, mathSum, mathIntegral,
toMathComponents

// Lint (не зависят от экземпляра)
lintChildren(children)
printLintResults(issues)

// Из docx
TableOfContents

// Стиль
DEFAULT_DOCX_STYLE, createDocxStyle
```

Методы экземпляра `createDocxGost()`:

```
run, paragraph, centered, paragraphWithMath,
h1, h2, h3, blank, pageBreak,
caption, tableCaption, createCaptionCounter,
placeholder, imageBlock, codeLine, codeBlock,
formulaMath, formulaInline,
makeTable,
makeTitlePage, makeContentSection, makeDocument, saveDocument,
diagramBlock,   ← рендер PlantUML + авторазмер + подпись
lint            ← проверка на пустые страницы и нарушения ГОСТ
```

### `@magomed-cmd/gost/plantuml`

```js
createPlantUmlGost(opts?)    // фабрика
autoImageSize(buffer, maxWidth?)
renderDiagram(source, opts?)
wrapWithSkin(source, dpi?, opts?)

// Builders (можно создавать напрямую)
UseCaseBuilder, ClassDiagramBuilder, SequenceBuilder, StateBuilder
```

Методы экземпляра `createPlantUmlGost()`:

```
useCaseDiagram()
classDiagram()
sequenceDiagram()
stateDiagram()
render(puml, opts?)
autoImageSize(buffer, maxWidth?)
```

### `@magomed-cmd/gost/excel`

```js
createExcelGost(opts?)
createExcelStyle(opts?)
```

Методы экземпляра `createExcelGost()`:

```
styleCell, setValue, setFormula,
applyBaseSheetSetup, mergeTitle, setFreeze,
applyZebraStriping, formatTableRange, getCellText, autoFitSheet,
addAhpMatrixBlock, addContentsSheet,
productFormula, absoluteRef, weightedSumFormula, listFormulaRange
```

---

## Частые ошибки

| Ошибка | Как правильно |
| --- | --- |
| `require("@magomed-cmd/gost")` | использовать subpath: `@magomed-cmd/gost/docx` |
| Подпись таблицы через `caption()` | `createCaptionCounter("Таблица", { table: true })` |
| Рисунок без ссылки в тексте | `paragraph(\`На рисунке ${figures.peek()} показан...\`)` до `imageBlock` |
| Формулы plain text | `formulaMath()` или `paragraphWithMath()` |
| Хардкод размеров `imageBlock(img, 400, 723)` | `autoImageSize(buffer)` |
| `pageBreak()` в конце `children` | пустая страница в конце — поймает `g.lint()` |
| Забыли `autoFitSheet()` | вызывать в конце каждого Excel-листа |
| `placeholder` в финальном отчёте | заменить на реальный `imageBlock` — `g.lint()` предупредит |

# @magomed-cmd/gost

Библиотека для программной генерации учебных отчётов по ГОСТ 7.32.

Форматы и возможности:
- `.docx` — заголовки, абзацы, формулы, таблицы, рисунки, листинги, титульник
- `.xlsx` — ячейки, формулы, стили, заморозка, МАИ
- UML-диаграммы через PlantUML прямо в коде, без внешних `.puml` файлов

Главная идея: фабрики `createDocxGost()`, `createExcelGost()`, `createPlantUmlGost()` создают независимые экземпляры со своими настройками стиля. Нет глобального состояния — один процесс может собирать несколько документов с разными параметрами.

---

## Быстрый выбор

| Задача | Использовать |
| --- | --- |
| Собрать Word-отчёт | `createDocxGost()` из `@magomed-cmd/gost/docx` |
| Абзацы, заголовки, формулы, таблицы | методы экземпляра `g` |
| Нумерация рисунков / таблиц | `g.createCaptionCounter("Рисунок")` |
| Вставить PlantUML диаграмму в DOCX | `g.diagramBlock(puml, counter, caption)` |
| Проверить на пустые страницы и нарушения ГОСТ | `g.lint(children)` перед `makeDocument` |
| Нарисовать UML-диаграмму кодом | `createPlantUmlGost()` из `@magomed-cmd/gost/plantuml` |
| Авторазмер PNG под ширину ГОСТ A4 | `autoImageSize(buffer)` |
| Сгенерировать Excel-файл | `createExcelGost()` из `@magomed-cmd/gost/excel` |
| Построить блок МАИ в Excel | `x.addAhpMatrixBlock(...)` |
| Переопределить шрифт/размер/цвета | опции фабрики |

---

## Установка и сборка

```bash
npm install
npm run build   # компиляция TypeScript → dist/
npm test        # 80 тестов
```

После сборки появляются файлы в `dist/`:

```
dist/docx-gost.js       dist/docx-gost.d.ts
dist/excel-gost.js      dist/excel-gost.d.ts
dist/plantuml-gost.js   dist/plantuml-gost.d.ts
dist/plantuml-render.js
dist/plantuml-size.js
dist/lint.js
```

Subpath entrypoints пакета:

```js
const { createDocxGost }     = require("@magomed-cmd/gost/docx");
const { createExcelGost }    = require("@magomed-cmd/gost/excel");
const { createPlantUmlGost } = require("@magomed-cmd/gost/plantuml");
```

Или через относительный путь:

```js
const { createDocxGost }     = require("../gost/dist/docx-gost");
const { createExcelGost }    = require("../gost/dist/excel-gost");
const { createPlantUmlGost } = require("../gost/dist/plantuml-gost");
```

В `package.json` нет верхнеуровневого `main` — импортировать только через subpath. Пакет рассчитан на Node.js 18+.

---

## DOCX

### Минимальный пример

```js
const path = require("path");
const fs   = require("fs");
const { createDocxGost, TableOfContents, mathFraction, mathSub } = require("@magomed-cmd/gost/docx");

const g = createDocxGost();
const {
  h1, h2, paragraph, paragraphWithMath, formulaMath,
  makeTable, imageBlock, createCaptionCounter,
  makeTitlePage, makeContentSection, makeDocument, pageBreak, saveDocument,
} = g;

const tables  = createCaptionCounter("Таблица", { table: true });
const figures = createCaptionCounter("Рисунок");

const children = [
  h1("Содержание"),
  new TableOfContents("Содержание", { hyperlink: false, headingStyleRange: "1-3" }),
  pageBreak(),

  h1("1 Расчётная часть"),
  h2("1.1 Исходные данные"),
  paragraph(`Исходная матрица приведена в таблице ${tables.peek()}.`),
  tables.caption("платёжная матрица"),
  makeTable([
    ["", "b1", "b2"],
    ["a1", 3, 5],
    ["a2", 4, 2],
  ], { widths: [2000, 3600, 3755] }),

  h2("1.2 Формулы"),
  paragraphWithMath([
    "Оптимальное значение находится как ",
    mathSub("V", "max"),
    " среди допустимых стратегий.",
  ]),
  formulaMath(["ν = ", mathFraction("48", "13"), " ≈ 3.692"]),

  paragraph(`Графическое решение показано на рисунке ${figures.peek()}.`),
  imageBlock(fs.readFileSync("graph.png"), 600, 420),
  figures.caption("графический метод"),
];

g.lint(children); // проверка перед сборкой

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

### Архитектура

`createDocxGost(options)` возвращает объект со всеми функциями генерации Word-документа. Все методы используют один общий стиль `style`, созданный из `DEFAULT_DOCX_STYLE` и переданных опций.

```js
const g = createDocxGost({
  style: {
    font: "Times New Roman",
    size: 28,
    sizeCode: 20,
    contentWidth: 9355,
  },
});
```

Размеры шрифтов в `docx` задаются в half-points: `28` = 14 pt, `24` = 12 pt, `20` = 10 pt.
Большинство размеров страницы, отступов и ширин задаются в twip/DXA: 1 inch = 1440 twip.

### Настройка стиля

#### `DEFAULT_DOCX_STYLE`

| Поле | Значение | Смысл |
| --- | --- | --- |
| `FONT` | `Times New Roman` | основной шрифт |
| `SIZE` | `28` | основной размер 14 pt |
| `SIZE_TITLE` | `32` | размер крупного текста титульника |
| `SIZE_HEADER` | `24` | размер верхней части титульника |
| `SIZE_CODE` | `24` | размер кода по умолчанию |
| `INDENT` | `709` | красная строка, ~1.25 см |
| `LINE` | `360` | межстрочный интервал |
| `CONTENT_WIDTH` | `9355` | ширина контента для таблиц в DXA |
| `CODE_FONT` | `Courier New` | моноширинный шрифт листингов |
| `TABLE_CELL_LINE` | `276` | межстрочный интервал в ячейках |
| `TABLE_CELL_MARGINS` | `{ top, bottom, left, right }` | внутренние поля ячеек |
| `HEADING_COLOR` | `000000` | цвет заголовков |
| `CAPTION_BEFORE/AFTER` | `60`, `60` | интервалы подписи рисунка |
| `TABLE_CAPTION_BEFORE/AFTER` | `120`, `60` | интервалы подписи таблицы |
| `IMAGE_BEFORE/AFTER` | `120`, `60` | интервалы вокруг рисунка |
| `FORMULA_BEFORE/AFTER` | `80`, `80` | интервалы вокруг формулы |
| `PAGE` | A4 + поля | параметры страницы |
| `BORDER_SINGLE` | single black | одиночная граница |
| `BORDERS_ALL` | single black везде | границы таблиц |
| `TITLE_TAB_STOPS` | tab stops титульника | позиции подписи/ФИО |

#### `createDocxStyle(opts)`

Создаёт полный объект стиля. Принимает как верхнерегистровые поля, так и удобные aliases.

```js
const { createDocxStyle } = require("@magomed-cmd/gost/docx");

const style = createDocxStyle({
  font: "Arial",
  size: 24,
  sizeCode: 20,
  line: 360,
  page: { margin: { left: 1701, right: 850 } },
});
```

Поддерживаемые aliases:

| Alias | Полное поле |
| --- | --- |
| `font` | `FONT` |
| `size` | `SIZE` |
| `sizeTitle` | `SIZE_TITLE` |
| `sizeHeader` | `SIZE_HEADER` |
| `sizeCode` | `SIZE_CODE` |
| `indent` | `INDENT` |
| `line` | `LINE` |
| `contentWidth` | `CONTENT_WIDTH` |
| `codeFont` | `CODE_FONT` |
| `page` | `PAGE` |

### Текстовые элементы

#### `run(text, opts)`

Создаёт `TextRun` для ручного использования внутри низкоуровневых `docx`-структур. Обычно напрямую не нужен — для обычных абзацев лучше `paragraph()`.

```js
const r = g.run("важно", { bold: true });
```

Опции: `size`, `bold`, `italics`, `allCaps`, `underline`, `color`, `font`.

#### `paragraph(content, opts)`

Обычный абзац. По умолчанию: выравнивание по ширине, красная строка, `keepLines: true`.

```js
paragraph("Текст абзаца.");
paragraph("Без красной строки.", { noIndent: true });
paragraph("По центру.", { align: "center", noIndent: true });
paragraph([
  "Метод ",
  { text: "Джонсона", bold: true },
  " применяется для упорядочивания работ.",
]);
```

`content` — строка, число или массив фрагментов `{ text, bold, italics, size, color, font }`.

Опции абзаца:

| Поле | Смысл |
| --- | --- |
| `align` | выравнивание: `left`, `center`, `right`, `both` |
| `noIndent` | убрать красную строку |
| `indent` | свой отступ первой строки |
| `line` | межстрочный интервал |
| `before` / `after` | интервалы перед/после абзаца |
| `bold`, `italics`, `allCaps`, `underline` | оформление текста |
| `size`, `color`, `font` | размер, цвет, шрифт |
| `keepNext` | держать следующий абзац на той же странице |
| `keepLines` | не разрывать строки; по умолчанию `true` |

#### `centered(content, opts)`

То же, что `paragraph()`, но по центру и без красной строки.

```js
centered("РЕФЕРАТ", { bold: true });
```

#### `paragraphWithMath(parts, opts)`

Абзац, где обычный текст и математические компоненты идут в одной строке.

```js
paragraphWithMath([
  "Пусть ",
  mathSub("a", "i"),
  " — стратегия игрока A, тогда значение определяется как ",
  mathSubSup("x", "i", "*"),
  ".",
]);
```

Правило: если выражение содержит `=`, `≥`, `≤`, `≈`, индексы, дроби, степени или стрелки — использовать math API, а не plain text.

### Заголовки и структура

#### `h1(text)` / `h2(text)` / `h3(text)`

Заголовки уровней 1–3. `h1` автоматически переводит текст в uppercase и центрирует. Все заголовки получают `HeadingLevel` и попадают в `TableOfContents`.

```js
h1("Введение");
h2("1.1 Постановка задачи");
h3("Критерий Вальда");
```

#### `blank(opts)` / `pageBreak()`

```js
blank();           // пустой абзац для вертикального промежутка
blank({ line: 240 });
pageBreak();       // принудительный разрыв страницы
```

### Подписи и счётчики

#### `createCaptionCounter(prefix, opts)`

Локальный счётчик подписей. Хранит состояние внутри объекта, без глобальных переменных.

```js
const figures = createCaptionCounter("Рисунок");
const tables  = createCaptionCounter("Таблица", { table: true });
```

| Метод | Что делает |
| --- | --- |
| `peek()` | возвращает текущий номер без инкремента |
| `caption(text)` | создаёт подпись и инкрементирует счётчик |

Для таблиц обязательно передавать `{ table: true }` — иначе подпись будет по центру, что неверно по ГОСТ.

Начать нумерацию не с 1:
```js
const figures = createCaptionCounter("Рисунок", { start: 4 });
```

Правильный паттерн:
```js
// Рисунок — ссылка, рисунок, подпись под рисунком
paragraph(`На рисунке ${figures.peek()} показан результат.`);
imageBlock(buffer, 600, 420);
figures.caption("результат запуска программы");

// Таблица — ссылка, подпись над таблицей, таблица
paragraph(`Данные приведены в таблице ${tables.peek()}.`);
tables.caption("исходные данные");
makeTable(rows);
```

#### `caption(text, opts)` / `tableCaption(text, opts)`

Прямые вызовы без счётчика. `caption` — по центру (для рисунков), `tableCaption` — по левому краю (для таблиц).

#### `placeholder(label, opts)`

Серый блок-заглушка когда изображение ещё не готово.

```js
placeholder("[Рисунок — скриншот главного окна]");
```

### Рисунки

#### `imageBlock(pathOrBuffer, width, height, opts)`

```js
imageBlock(fs.readFileSync("graph.png"), 600, 420);
imageBlock(fs.readFileSync("photo.jpg"), 400, 300, { imageType: "jpg" });
```

По умолчанию `imageType: "png"`. Поддерживаются `"jpg"`, `"gif"`, `"bmp"`.

По умолчанию `keepNext: true` — изображение держится со следующим абзацем (подписью), что предотвращает разрыв страницы между рисунком и подписью.

#### `diagramBlock(pumlSource, counter, captionText, opts?)`

Комбо-метод: рендерит PlantUML строку → авторазмер → `imageBlock` + подпись. Спредится в `children`.

```js
const { createPlantUmlGost } = require("@magomed-cmd/gost/plantuml");
const p = createPlantUmlGost();

const children = [
  paragraph(`На рисунке ${figures.peek()} представлена диаграмма классов.`),
  ...await g.diagramBlock(
    p.classDiagram()
      .interface("RepairStrategy", ["+ getType(): String"])
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
| `dpi` | 150 | разрешение PNG (300 для финального отчёта на печать) |
| `maxWidth` | 624 | максимальная ширина px (ГОСТ A4) |
| `skinparams` | — | переопределение skinparam PlantUML |

Если диаграмма выше 900px — в консоль выводится предупреждение.

### Листинги кода

#### `codeLine(text, opts)`

Одна строка кода без красной строки, моноширинным шрифтом.

```js
codeLine("npm run build");
codeLine("int main() { return 0; }", { size: 20 });
```

#### `codeBlock(lines, opts)`

Массив строк кода. Раскрывается через spread.

```js
const source = fs.readFileSync("program.cpp", "utf8").split("\n");

const children = [
  h1("Исходный код"),
  ...codeBlock(source),
  ...codeBlock(source, { size: 20 }), // 10 pt для компактного листинга
];
```

### Таблицы

#### `makeTable(rows, opts)`

```js
makeTable([
  ["Критерий", "Значение"],
  ["MM", "E1"],
  [{ text: "Итого", bold: true }, { text: 23, bold: true }],
], { widths: [4000, 5355] });
```

Правила реализации:
- ширина задаётся только через `WidthType.DXA`
- первая строка по умолчанию считается заголовком (жирный + центр)
- ширины колонок задавать явно через `widths`

Опции таблицы:

| Поле | Смысл |
| --- | --- |
| `widths` | массив ширин колонок в DXA |
| `headerRow` | индекс строки-заголовка, по умолчанию `0` |
| `align` | выравнивание обычных ячеек |
| `centerHeader` | центрировать заголовок, по умолчанию `true` |
| `fontSize` | размер текста в ячейках |
| `cellLine` | межстрочный интервал в ячейках |
| `contentWidth` | ширина таблицы вместо `style.CONTENT_WIDTH` |
| `borders` | объект границ `docx` |

Объект ячейки:

| Поле | Смысл |
| --- | --- |
| `text` | значение: строка, число, массив фрагментов или math-компонент |
| `align` | выравнивание |
| `bold`, `italics` | оформление |
| `size`, `color`, `font` | шрифт |
| `colspan`, `rowspan` | объединение ячеек |
| `shading` | заливка в формате `docx` |

### Математика

Математические функции экспортируются отдельно из модуля — они не зависят от экземпляра.

```js
const {
  createDocxGost, mathFraction, mathSub, mathSup, mathSubSup,
  mathParen, mathBracket, mathBrace, mathRoot, mathSum, mathIntegral,
} = require("@magomed-cmd/gost/docx");
```

#### `formulaMath(content, opts)`

Отдельная центрированная формула.

```js
formulaMath(["Cmax = 23"]);
formulaMath(["ν = ", mathFraction("48", "13"), " ≈ 3.692"]);
formulaMath([mathSubSup("x", "i", "*"), " = 0"]);
formulaMath([mathSum(mathSub("x", "i"), { subScript: "i=1", superScript: "n" })]);
formulaMath([mathIntegral("f(x)dx", { subScript: "0", superScript: "1" })]);
```

#### `formulaInline(label, mathContent, opts)`

Абзац вида `label + формула`.

```js
formulaInline("Получаем:", ["q* = ", mathFraction("10", "13")]);
```

#### Math-хелперы

| Функция | Смысл |
| --- | --- |
| `mathFraction(num, den)` | дробь |
| `mathSub(base, sub)` | нижний индекс |
| `mathSup(base, sup)` | верхний индекс |
| `mathSubSup(base, sub, sup)` | оба индекса |
| `mathParen(...parts)` | круглые скобки |
| `mathBracket(...parts)` | квадратные скобки |
| `mathBrace(...parts)` | фигурные скобки |
| `mathRoot(content, degree?)` | корень |
| `mathSum(content, opts)` | сумма с пределами |
| `mathIntegral(content, opts)` | интеграл с пределами |
| `mathExpr(...parts)` | произвольное выражение |

### Титульник, секции и документ

#### `makeTitlePage(opts)`

Создаёт первую секцию с титульным листом и нижним колонтитулом `Таганрог {year}`.

```js
makeTitlePage({
  workType: "ОТЧЁТ",
  subtitle: ["о выполнении лабораторной работы №4", "на тему: «TramLine»"],
  discipline: "по дисциплине: «Архитектура и технологии разработки ПО»",
  group: "КТбо3-7",
  author: "Валиев М. М.",
  teacher: "Лутай В. Н.",
  teacherTitle: "Доцент кафедры",
  year: 2026,
});
```

| Поле | Обязательное | Смысл |
| --- | --- | --- |
| `workType` | да | тип работы |
| `subtitle` | нет | строка или массив строк под названием |
| `discipline` | да | дисциплина |
| `group` | да | группа |
| `author` | да | ФИО студента |
| `teacher` | да | ФИО преподавателя |
| `teacherTitle` | нет | должность, по умолчанию `Доцент кафедры` |
| `year` | да | год в нижнем колонтитуле |

#### `makeContentSection(children)`

Создаёт секцию основного содержания с нумерацией страниц в нижнем колонтитуле.

#### `makeDocument(sections, opts?)`

Создаёт `docx.Document` из `DocxSection[]`. По умолчанию `settings.updateFields = true` — `TableOfContents` обновляется Word при открытии.

```js
const doc = makeDocument([
  makeTitlePage(titleOpts),
  makeContentSection(children),
]);
```

#### `saveDocument(doc, outputPath)`

```js
await saveDocument(doc, path.join(__dirname, "report.docx"));
```

#### `TableOfContents`

```js
const { TableOfContents } = require("@magomed-cmd/gost/docx");

new TableOfContents("Содержание", { hyperlink: false, headingStyleRange: "1-3" });
```

### Lint — проверка документа

`g.lint(children)` анализирует массив **до** сборки документа, выводит предупреждения в консоль и возвращает `LintIssue[]`.

```js
g.lint(children);
const doc = makeDocument([makeTitlePage(opts), makeContentSection(children)]);
```

Пример вывода:
```
  ⚠ [WARN] pos  14  Два pageBreak подряд на позиции 14–15 — пустая страница.
  ⚠ [WARN] pos  47  image на позиции 47 отделён от подписи явным pageBreak.
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

Для использования в CI:

```js
const { lintChildren, printLintResults } = require("@magomed-cmd/gost/docx");

const issues = lintChildren(children);
printLintResults(issues);
if (issues.filter(i => i.level === "warn").length > 0) process.exit(1);
```

### Рекомендуемый шаблон

```js
const g = createDocxGost();
const {
  h1, h2, paragraph, makeTable, makeTitlePage,
  makeContentSection, makeDocument, saveDocument,
  createCaptionCounter, pageBreak,
} = g;

const tables  = createCaptionCounter("Таблица", { table: true });
const figures = createCaptionCounter("Рисунок");

const children = [
  h1("Содержание"),
  new TableOfContents("Содержание", { hyperlink: false, headingStyleRange: "1-3" }),
  pageBreak(),

  h1("Введение"),
  paragraph("Во введении описывается цель работы."),
  pageBreak(),

  h1("1 Теоретическая часть"),
  h2("1.1 Исходные данные"),
  paragraph(`Исходные данные представлены в таблице ${tables.peek()}.`),
  tables.caption("исходные данные"),
  makeTable(rows),

  pageBreak(),
  h1("Заключение"),
  paragraph("В заключении приводятся основные результаты."),
];

g.lint(children);
await saveDocument(
  makeDocument([makeTitlePage(titleOpts), makeContentSection(children)]),
  outputPath
);
```

### Частые ошибки

| Ошибка | Как правильно |
| --- | --- |
| Подпись таблицы через `caption()` | `createCaptionCounter("Таблица", { table: true })` |
| Рисунок без ссылки в тексте | `paragraph(\`На рисунке ${figures.peek()} показан...\`)` до `imageBlock` |
| Таблица без ссылки в тексте | `paragraph(\`... в таблице ${tables.peek()}.\`)` до `tables.caption()` |
| Математика plain text в `paragraph()` | `formulaMath()` или `paragraphWithMath()` |
| Хардкод размеров `imageBlock(img, 400, 723)` | `autoImageSize(buffer)` из plantuml-модуля |
| `pageBreak()` в конце `children` | пустая страница в конце — поймает `g.lint()` |
| Глобальные счётчики | локальные `const figures = createCaptionCounter(...)` внутри функции сборки |
| `placeholder` в финальном отчёте | заменить на `imageBlock` — `g.lint()` предупредит |

---

## PlantUML

Модуль позволяет писать PlantUML прямо в скрипте — никаких внешних `.puml` файлов. Диаграмма рендерится на лету и вставляется в DOCX через `g.diagramBlock()` или вручную.

Требует установленной Java.

```js
const { createPlantUmlGost } = require("@magomed-cmd/gost/plantuml");

const p = createPlantUmlGost();       // dpi: 150 по умолчанию
const p = createPlantUmlGost({ dpi: 300 }); // для финального отчёта
```

### Builders

Все builders поддерживают `.add(line)` как escape hatch для любой строки PlantUML которую builder не покрывает, и `.blank()` для визуального разделения групп.

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
  .extend("UC1", "UC3")
  .package("Система", b => b.usecase("Внутреннее действие"))
  .build()
```

Методы: `actor`, `usecase`, `arrow`, `extend`, `include`, `package`.

#### `p.classDiagram()`

`build()` автоматически добавляет `hide circle`.

```js
p.classDiagram()
  .interface("RepairStrategy", ["+ getType(): String", "+ getRepairTimeMs(): int"])
  .class("ElectricianRepair", ["+ getType(): String"])
  .abstract("BaseRepair", ["# timeMs: int"])
  .enum("TramState", ["ON_LINE", "IN_DEPOT", "BROKEN", "RESERVE"])
  .implements("ElectricianRepair", "RepairStrategy")
  .extends("ElectricianRepair", "BaseRepair")
  .association("Dispatcher", "TramRoute", "управляет")
  .composition("TramRoute", "Tram")
  .aggregation("Route", "Stop")
  .dependency("Dispatcher", "RepairStrategy", "использует")
  .package("tramline.model", b => b.class("Tram").class("TramRoute"))
  .build()
```

Методы: `class`, `abstract`, `interface`, `enum`, `extends`, `implements`, `association`, `dependency`, `composition`, `aggregation`, `package`.

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
  .return("D", "UI", "publish(BREAKDOWN)")
  .deactivate("D")
  .note("Всё в dispatcher-thread", "right", "D")
  .group("alt резерв есть", b => b.arrow("D", "D", "deployReserve()"))
  .build()
```

Методы: `participant`, `actor`, `arrow`, `return`, `activate`, `deactivate`, `note`, `divider`, `group`.

#### `p.stateDiagram()`

`transition` принимает опциональное направление `"down"`, `"up"`, `"left"`, `"right"` — помогает избежать пересечений стрелок.

```js
p.stateDiagram()
  .initial("ON_LINE")
  .state("ON_LINE", "На линии")
  .state("BROKEN",  "Сломан")
  .state("IN_DEPOT","В парке")
  .state("RESERVE", "Резерв")
  .transition("ON_LINE", "BROKEN",   "поломка",        "down")
  .transition("BROKEN",  "IN_DEPOT", "отремонтирован")
  .transition("IN_DEPOT","ON_LINE",  "выход на линию")
  .transition("ON_LINE", "RESERVE",  "избыток")
  .transition("RESERVE", "ON_LINE",  "нехватка")
  .final("IN_DEPOT")
  .build()
```

Методы: `state`, `transition`, `initial`, `final`.

### `p.render(puml, opts?)`

Рендерит puml-строку, возвращает PNG Buffer. Базовые skinparam добавляются автоматически.

```js
const png = await p.render(p.classDiagram().class("Foo").build());
const png = await p.render(puml, { dpi: 300, format: "png" });
```

### `autoImageSize(buffer, maxWidth?)`

Читает реальные размеры PNG из буфера и масштабирует пропорционально. Максимальная ширина ГОСТ A4 = **624px**.

```js
const { autoImageSize } = require("@magomed-cmd/gost/plantuml");

const png = await p.render(puml);
const { width, height } = autoImageSize(png);       // maxWidth = 624
const { width, height } = autoImageSize(png, 400);  // своя ширина

imageBlock(png, width, height);
```

### `wrapWithSkin(source, dpi?, opts?)`

Добавляет базовые skinparam к puml-строке. Вызывается внутри `render()` автоматически. Базовые значения: Times New Roman 12pt, Shadowing false, nodesep 60, ranksep 50.

```js
const { wrapWithSkin } = require("@magomed-cmd/gost/plantuml");

const styled = wrapWithSkin(rawPuml, 300, { skinparams: { backgroundColor: "white" } });
```

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

`createExcelGost(opts)` принимает поля в верхнем регистре и удобные aliases. Если переданы оба — верхний регистр имеет приоритет.

```js
const x = createExcelGost({
  fontSize: 14,          // FONT_SIZE
  fontName: "Arial",     // FONT_NAME
  headerFill: "FFF2CC",  // HEADER_FILL
  resultFill: "D9EAD3",  // RESULT_FILL
  tabColor: "FFF2CC",    // TAB_COLOR
});
```

Поля стиля:

| Поле | По умолчанию | Смысл |
| --- | --- | --- |
| `HEADER_FILL` | `FFF2CC` | заливка заголовков |
| `RESULT_FILL` | `D9EAD3` | заливка итоговых ячеек |
| `ALT_ROW_FILL` | `FFF9F0` | заливка чётных строк |
| `REGRET_FILL` | `FCE4D6` | заливка для матриц сожалений |
| `FONT_NAME` | `Times New Roman` | основной шрифт |
| `FONT_SIZE` | `14` | основной размер |
| `TAB_COLOR` | `FFF2CC` | цвет вкладки листа |
| `DEFAULT_ROW_HEIGHT` | `20` | высота строки по умолчанию |
| `PAGE_ORIENTATION` | `landscape` | ориентация страницы |
| `MAX_COLUMN_WIDTH` | `80` | максимум ширины при autofit |

### Стили ячеек

#### `styleCell(cell, opts)`

```js
x.styleCell(sheet.getCell("A1"), {
  bold: true, fill: x.style.HEADER_FILL,
  border: true, align: "center",
});
```

Опции: `fontName`, `size`, `bold`, `color`, `align`, `wrapText`, `fill`, `border`, `numFmt`.

#### `setValue(sheet, address, value, style?)` / `setFormula(sheet, address, formula, style?)`

```js
x.setValue(sheet, "A1", "Показатель", x.HEADER_STYLE);
x.setValue(sheet, "B1", 12.345, { border: true, numFmt: "0.000" });
x.setFormula(sheet, "B10", "SUM(B4:B9)", { border: true, numFmt: "0.0000" });
```

`setFormula` принимает формулу без начального `=`.

#### Готовые стили экземпляра

```js
x.HEADER_STYLE // { bold: true, border: true, fill: HEADER_FILL, align: "center" }
x.RESULT_STYLE // { bold: true, border: "medium", fill: RESULT_FILL, align: "center" }
```

Для меток слева:
```js
const LABEL_STYLE = { bold: true, border: true, fill: x.style.HEADER_FILL, align: "left" };
```

### Настройка листа

#### `applyBaseSheetSetup(sheet, widths)`

Применяет базовые параметры: высота строк, ориентация, fit-to-page, ширины колонок.

```js
x.applyBaseSheetSetup(sheet, [18, 14, 14, 14, 20]);
```

#### `mergeTitle(sheet, range, text)`

Объединяет диапазон и пишет крупный заголовок.

```js
x.mergeTitle(sheet, "A1:F1", "РАСЧЁТ КРИТЕРИЕВ");
```

#### `setFreeze(sheet, ySplit, activeCell)`

Замораживает верхние строки через `sheet.views`.

```js
x.setFreeze(sheet, 3, "A4");
```

#### `applyZebraStriping(sheet, startRow, endRow, startCol, endCol, headerRows)`

Применяет мягкую заливку чётных строк. Строки из `headerRows` пропускаются.

```js
x.applyZebraStriping(sheet, 3, 20, 1, 8, new Set([3]));
```

#### `formatTableRange(sheet, startRow, endRow, startCol, endCol, headerRow)`

Оформляет прямоугольный диапазон как таблицу.

```js
x.formatTableRange(sheet, 3, 12, 1, 6, 3);
```

#### `autoFitSheet(sheet)`

Автоматически подбирает ширину колонок и высоту строк. Вызывать в конце формирования листа.

```js
x.autoFitSheet(sheet);
```

### Формульные хелперы

Доступны только как методы экземпляра.

```js
x.productFormula(["B4", "C4", "D4"], 3);            // (B4*C4*D4)^(1/3)
x.absoluteRef("B12");                                // $B$12
x.weightedSumFormula(["B2", "C2"], ["$B$1", "$C$1"]); // B2*$B$1+C2*$C$1
x.listFormulaRange("МАИ", ["B2", "B3"]);             // 'МАИ'!B2,'МАИ'!B3
```

### Блок МАИ

#### `addAhpMatrixBlock(sheet, startRow, title, labels, matrix, randomIndex)`

Строит полный блок матрицы парных сравнений: матрица, геометрические средние, веса, λ_max, ИС, ОС.

```js
const result = x.addAhpMatrixBlock(
  sheet, 5,
  "Матрица парных сравнений критериев",
  ["К1", "К2", "К3"],
  [[1, 3, 5], [1/3, 1, 2], [1/5, 1/2, 1]],
  0.58
);
// result: { weightColumn, weightStartRow, weightEndRow, osCell, nextRow }
```

`labels` и `matrix` валидируются: при несоответствии размеров бросается `RangeError` до записи данных. Колонки поддерживаются после `Z` (`AA`, `AB`, ...).

### Содержание workbook

#### `addContentsSheet(workbook, sheetNames)`

Создаёт лист «Содержание» со ссылками на остальные листы.

```js
x.addContentsSheet(workbook, ["Матрица", "Критерии", "МАИ"]);
```

### Рекомендуемый шаблон

```js
const ExcelJS = require("exceljs");
const { createExcelGost } = require("@magomed-cmd/gost/excel");

async function buildWorkbook(outputPath) {
  const workbook = new ExcelJS.Workbook();
  const x = createExcelGost({ fontSize: 14 });

  const sheet = workbook.addWorksheet("Расчёт", {
    properties: { tabColor: { argb: x.style.TAB_COLOR } },
  });

  x.applyBaseSheetSetup(sheet, [24, 18, 18, 18]);
  x.mergeTitle(sheet, "A1:D1", "РАСЧЁТ");

  const LABEL_STYLE = { bold: true, border: true, fill: x.style.HEADER_FILL, align: "left" };
  x.setValue(sheet, "A3", "N=", LABEL_STYLE);
  x.setValue(sheet, "B3", 7, { border: true });
  x.setValue(sheet, "A5", "Показатель", x.HEADER_STYLE);
  x.setValue(sheet, "B5", "Значение", x.HEADER_STYLE);
  x.setFormula(sheet, "B6", "SUM(B3:B3)", { border: true, numFmt: "0.000" });

  x.formatTableRange(sheet, 5, 6, 1, 2, 5);
  x.setFreeze(sheet, 4, "A5");
  x.autoFitSheet(sheet);

  await workbook.xlsx.writeFile(outputPath);
}
```

### Частые ошибки

| Ошибка | Как правильно |
| --- | --- |
| Метки слева без заливки и границ | `{ bold: true, border: true, fill: x.style.HEADER_FILL, align: "left" }` |
| Формула с начальным `=` | `setFormula(sheet, "A1", "SUM(B1:B2)")` |
| Забыли `autoFitSheet()` | вызывать в конце каждого листа |
| Нечитаемый мелкий шрифт | `createExcelGost({ fontSize: 14 })` |
| Workbook создаётся библиотекой | workbook создаёт пользователь: `new ExcelJS.Workbook()` |
| Excel-хелперы импортируются напрямую | использовать `x.productFormula()` и другие методы экземпляра |

---

## Тесты

```bash
npm test               # все тесты
npm run test:coverage  # с отчётом покрытия
```

80 тестов в 5 файлах:

| Файл | Что покрывается |
| --- | --- |
| `plantuml-size.test.ts` | `autoImageSize` — масштабирование, пропорции, edge cases |
| `plantuml-render.test.ts` | `wrapWithSkin` — skinparam injection, переопределение, структура |
| `plantuml-builders.test.ts` | все 4 builder-а и все их методы |
| `docx-gost.test.ts` | публичный API, `CaptionCounter`, `makeTable`, `makeDocument` |
| `lint.test.ts` | все 11 lint-правил |

---

## Соглашения

1. Все DOCX-элементы класть в массив `children`.
2. Заголовки делать через `h1`, `h2`, `h3`, а не жирными абзацами.
3. Таблицу всегда предварять абзацем-ссылкой и `tables.caption()` над таблицей.
4. Для таблиц создавать счётчик только так: `createCaptionCounter("Таблица", { table: true })`.
5. Рисунок всегда: ссылка → `imageBlock()` или `placeholder()` → `figures.caption()`.
6. Формулы с индексами, дробями, знаками сравнения — через `formulaMath()` или `paragraphWithMath()`.
7. Исходный код больших программ помещать в приложение через `codeBlock()`.
8. Вызывать `g.lint(children)` перед `makeDocument()`.
9. В конце каждого Excel-листа вызывать `autoFitSheet()`.
10. Не хранить ФИО, группы, номера вариантов внутри библиотеки — передавать в скрипт сборки.
11. Не добавлять side effects при импорте — генерация запускается только явным вызовом.

---

## Экспорты

### DOCX (`@magomed-cmd/gost/docx`)

Из модуля:

```js
createDocxGost, DEFAULT_DOCX_STYLE, createDocxStyle,
lintChildren, printLintResults,
mathExpr, mathFraction, mathSub, mathSup, mathSubSup,
mathParen, mathBracket, mathBrace, mathRoot, mathSum, mathIntegral,
toMathComponents,
TableOfContents,
```

Методы экземпляра:

```js
style,
run, paragraph, centered, paragraphWithMath,
h1, h2, h3, blank, pageBreak,
caption, tableCaption, createCaptionCounter,
placeholder, imageBlock, codeLine, codeBlock,
formulaMath, formulaInline,
makeTable,
makeTitlePage, makeContentSection, makeDocument, saveDocument,
diagramBlock,
lint,
```

### PlantUML (`@magomed-cmd/gost/plantuml`)

Из модуля:

```js
createPlantUmlGost,
autoImageSize, renderDiagram, wrapWithSkin,
UseCaseBuilder, ClassDiagramBuilder, SequenceBuilder, StateBuilder,
```

Методы экземпляра:

```js
useCaseDiagram, classDiagram, sequenceDiagram, stateDiagram,
render, autoImageSize,
```

### Excel (`@magomed-cmd/gost/excel`)

Из модуля:

```js
createExcelGost, createExcelStyle,
```

Методы экземпляра:

```js
style, HEADER_STYLE, RESULT_STYLE,
styleCell, setValue, setFormula,
applyBaseSheetSetup, mergeTitle, setFreeze,
applyZebraStriping, formatTableRange, getCellText, autoFitSheet,
addAhpMatrixBlock, addContentsSheet,
productFormula, absoluteRef, weightedSumFormula, listFormulaRange,
```

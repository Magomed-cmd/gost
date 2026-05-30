# GOST report libraries

Локальная библиотека для генерации отчетов по ГОСТ 7.32 в двух форматах:

- `.docx` через `docx`;
- `.xlsx` через `exceljs`.

Библиотека рассчитана на программную сборку учебных отчетов: титульный лист, содержание, заголовки, абзацы, формулы, таблицы, рисунки, листинги кода, Excel-листы с единым стилем и формулами.

Главная идея новой TypeScript-версии: использовать фабрики `createDocxGost()` и `createExcelGost()`. Фабрика создает независимый экземпляр библиотеки со своими настройками стиля. Это важно: нет глобального изменяемого состояния, один Node.js-процесс может собирать несколько документов с разными шрифтами, размерами и цветами.

## Быстрый выбор API

| Задача | Использовать |
| --- | --- |
| Собрать Word-отчет | `createDocxGost()` из `@codex/gost/docx` |
| Сделать абзацы, заголовки, формулы, таблицы | методы экземпляра `docxGost` |
| Сделать нумерацию рисунков | `createCaptionCounter("Рисунок")` |
| Сделать нумерацию таблиц по ГОСТ | `createCaptionCounter("Таблица", { table: true })` |
| Сгенерировать Excel-файл | `createExcelGost()` из `@codex/gost/excel` |
| Заполнить и оформить ячейку Excel | `setValue()`, `setFormula()`, `styleCell()` |
| Построить блок МАИ в Excel | `addAhpMatrixBlock()` |
| Переопределить шрифт/размер/цвета | опции фабрики |

## Установка и сборка

Внутри папки `lib`:

```bash
npm install
npm run build
```

После сборки появляются CommonJS-файлы:

```text
dist/docx-gost.js
dist/docx-gost.d.ts
dist/excel-gost.js
dist/excel-gost.d.ts
```

Публичные entrypoints пакета:

```js
const { createDocxGost } = require("@codex/gost/docx");
const { createExcelGost } = require("@codex/gost/excel");
```

В `package.json` нет верхнеуровневого `main`, поэтому импортировать нужно через subpath `@codex/gost/docx` или `@codex/gost/excel`. Пакет рассчитан на Node.js 18+ и публикует только `dist` и `README.md`.

В проекте также остаются старые CommonJS-файлы `docx-gost.js` и `excel-gost.js`. Они нужны для совместимости старых скриптов. Для нового кода предпочтителен factory API из package subpaths.

## DOCX: минимальный пример

```js
const path = require("path");
const fs = require("fs");
const {
  createDocxGost,
  TableOfContents,
  mathFraction,
  mathSub,
} = require("@codex/gost/docx");

const g = createDocxGost();

const {
  h1,
  h2,
  paragraph,
  paragraphWithMath,
  formulaMath,
  makeTable,
  imageBlock,
  createCaptionCounter,
  makeTitlePage,
  makeContentSection,
  makeDocument,
  pageBreak,
  saveDocument,
} = g;

const tables = createCaptionCounter("Таблица", { table: true });
const figures = createCaptionCounter("Рисунок");

const graphPath = path.join(__dirname, "..", "gen", "graph.png");

const children = [
  h1("Содержание"),
  new TableOfContents("Содержание", { hyperlink: false, headingStyleRange: "1-3" }),
  pageBreak(),

  h1("1 Расчетная часть"),
  h2("1.1 Исходные данные"),
  paragraph(`Исходная матрица приведена в таблице ${tables.peek()}.`),
  tables.caption("платежная матрица"),
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
  imageBlock(fs.readFileSync(graphPath), 600, 420),
  figures.caption("графический метод"),
];

const doc = makeDocument([
  makeTitlePage({
    workType: "ИНДИВИДУАЛЬНОЕ ЗАДАНИЕ",
    subtitle: "по теории принятия решений",
    discipline: "по дисциплине: «Теория принятия решений»",
    group: "КТбо3-7",
    author: "Бораев Р. Х.",
    teacher: "Родзин С. И.",
    year: 2026,
  }),
  makeContentSection(children),
]);

saveDocument(doc, path.join(__dirname, "..", "gen", "report.docx"));
```

## DOCX: архитектура

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

Размеры шрифтов в `docx` задаются в half-points: `28` означает 14 pt, `24` означает 12 pt, `20` означает 10 pt.

Большинство размеров страницы, отступов, ширин таблиц и межстрочных интервалов задаются в twip/DXA. Это стандартные единицы Word: 1 inch = 1440 twip.

## DOCX: настройка стиля

### `DEFAULT_DOCX_STYLE`

Базовые значения:

| Поле | Значение | Смысл |
| --- | --- | --- |
| `FONT` | `Times New Roman` | основной шрифт |
| `SIZE` | `28` | основной размер 14 pt |
| `SIZE_TITLE` | `32` | размер крупного текста титульного листа |
| `SIZE_HEADER` | `24` | размер верхней части титульного листа |
| `SIZE_CODE` | `24` | размер кода по умолчанию |
| `INDENT` | `709` | красная строка, примерно 1.25 см |
| `LINE` | `360` | межстрочный интервал |
| `CONTENT_WIDTH` | `9355` | ширина контента для таблиц в DXA |
| `CODE_FONT` | `Courier New` | моноширинный шрифт листингов |
| `TABLE_CELL_LINE` | `276` | межстрочный интервал в ячейках таблиц |
| `TABLE_CELL_MARGINS` | `{ top, bottom, left, right }` | внутренние поля ячеек |
| `HEADING_COLOR` | `000000` | цвет заголовков |
| `CAPTION_BEFORE`, `CAPTION_AFTER` | `60`, `60` | интервалы подписи рисунка |
| `TABLE_CAPTION_BEFORE`, `TABLE_CAPTION_AFTER` | `120`, `60` | интервалы подписи таблицы |
| `IMAGE_BEFORE`, `IMAGE_AFTER` | `120`, `60` | интервалы вокруг рисунка |
| `FORMULA_BEFORE`, `FORMULA_AFTER` | `80`, `80` | интервалы вокруг отдельной формулы |
| `PAGE` | A4 + поля | параметры страницы |
| `BORDER_SINGLE` | single black | одиночная граница |
| `BORDERS_ALL` | single black на всех сторонах | границы таблиц |
| `TITLE_TAB_STOPS` | tab stops титульника | позиции подписи/ФИО |

### `createDocxStyle(opts)`

Создает полный объект стиля. Можно передавать как верхнерегистровые поля, так и удобные alias-поля.

```js
const { createDocxStyle } = require("@codex/gost/docx");

const style = createDocxStyle({
  font: "Arial",
  size: 24,
  sizeCode: 20,
  line: 360,
  page: {
    margin: { left: 1701, right: 850 },
  },
});
```

Поддерживаемые alias:

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

Если нужно изменить только один параметр, не надо копировать весь стиль.

```js
const g = createDocxGost({ style: { sizeCode: 20 } });
```

## DOCX: текстовые элементы

### `run(text, opts)`

Создает `TextRun` для ручного использования внутри низкоуровневых `docx`-структур.

```js
const r = g.run("важно", { bold: true });
```

Опции:

| Поле | Смысл |
| --- | --- |
| `size` | размер в half-points |
| `bold` | жирный |
| `italics` | курсив |
| `allCaps` | все заглавные |
| `underline` | подчеркивание: `true` или объект `docx` |
| `color` | цвет без `#`, например `FF0000` |
| `font` | шрифт |

Обычно напрямую `run()` не нужен. Для обычных абзацев лучше использовать `paragraph()`.

### `paragraph(content, opts)`

Создает обычный абзац. По умолчанию:

- выравнивание по ширине;
- первая строка с отступом;
- `keepLines: true`;
- шрифт и размер из стиля экземпляра.

```js
paragraph("Текст абзаца.");
paragraph("Без красной строки.", { noIndent: true });
paragraph("По центру.", { align: "center", noIndent: true });
paragraph("Жирный текст.", { bold: true });
```

`content` может быть строкой, числом или массивом фрагментов:

```js
paragraph([
  "Метод ",
  { text: "Джонсона", bold: true },
  " применяется для упорядочивания работ.",
]);
```

Опции абзаца:

| Поле | Смысл |
| --- | --- |
| `align` | выравнивание: `left`, `center`, `right`, `both`/значение `AlignmentType` |
| `noIndent` | убрать красную строку |
| `indent` | свой отступ первой строки |
| `line` | межстрочный интервал |
| `before` | интервал перед абзацем |
| `after` | интервал после абзаца |
| `bold`, `italics`, `allCaps`, `underline` | оформление текста |
| `size`, `color`, `font` | размер, цвет, шрифт |
| `keepNext` | держать следующий абзац на той же странице |
| `keepLines` | не разрывать строки абзаца; по умолчанию `true` |
| `imageType` | тип изображения для `imageBlock()`: `png`, `jpg`, `gif`, `bmp` |

### `centered(content, opts)`

То же, что `paragraph()`, но с выравниванием по центру и без красной строки.

```js
centered("РЕФЕРАТ", { bold: true });
```

### `paragraphWithMath(parts, opts)`

Создает абзац, где обычный текст и математические компоненты идут в одной строке. Использовать, когда формула является частью предложения.

```js
paragraphWithMath([
  "Пусть ",
  mathSub("a", "i"),
  " - стратегия игрока A, тогда значение определяется как ",
  mathSubSup("x", "i", "*"),
  ".",
]);
```

Правило для отчетов: если выражение содержит `=`, `≥`, `≤`, `≈`, индексы, дроби, степени или стрелки, его лучше оформлять через математический API, а не как plain text внутри `paragraph()`.

## DOCX: заголовки и структура

### `h1(text, opts)`

Заголовок первого уровня. Текст автоматически переводится в uppercase, центрируется и получает `HeadingLevel.HEADING_1`, поэтому попадает в содержание.

```js
h1("Введение");
h1("1 Практическая часть");
```

### `h2(text, opts)`

Заголовок второго уровня. Использовать для подразделов.

```js
h2("1.1 Постановка задачи");
```

### `h3(text, opts)`

Заголовок третьего уровня. Использовать для вложенных пунктов внутри `h2`.

```js
h3("Критерий Вальда");
```

### `blank(opts)`

Пустой абзац для вертикального промежутка.

```js
blank();
blank({ line: 240 });
```

### `pageBreak()`

Абзац с разрывом страницы.

```js
pageBreak();
```

## DOCX: подписи, рисунки и плейсхолдеры

### `caption(text, opts)`

Подпись по центру. Использовать для рисунков, схем и графиков. По ГОСТ подпись рисунка ставится под рисунком.

```js
imageBlock(buffer, 600, 420);
caption("Рисунок 1 - графический метод");
```

### `tableCaption(text, opts)`

Подпись таблицы. Выравнивание по левому краю, без красной строки. По ГОСТ подпись таблицы ставится над таблицей.

```js
tableCaption("Таблица 1 - исходная матрица");
makeTable(rows);
```

### `createCaptionCounter(prefix, opts)`

Создает локальный счетчик подписей. Счетчик хранит состояние внутри объекта, а не в глобальной переменной.

```js
const figures = createCaptionCounter("Рисунок");
const tables = createCaptionCounter("Таблица", { table: true });
```

Методы:

| Метод | Что делает |
| --- | --- |
| `peek()` | возвращает следующий номер без увеличения счетчика |
| `caption(text)` | создает подпись и увеличивает счетчик |

Для ссылок в тексте использовать `peek()`:

```js
paragraph(`Исходные данные приведены в таблице ${tables.peek()}.`);
tables.caption("исходная матрица");
makeTable(rows);

paragraph(`На рисунке ${figures.peek()} показана нижняя огибающая.`);
imageBlock(buffer, 600, 420);
figures.caption("графический метод");
```

Важно: для таблиц обязательно передавать `{ table: true }`, иначе счетчик будет использовать `caption()` и подпись получится по центру, что неверно для таблиц.

Можно начать нумерацию не с 1:

```js
const figures = createCaptionCounter("Рисунок", { start: 4 });
```

### `placeholder(label, opts)`

Создает серый блок-заглушку. Использовать, когда рисунок или скриншот еще не готов, но в отчете нужно оставить место.

```js
placeholder("[Рисунок 2 - запуск программы]");
```

Если реальное изображение уже есть, использовать `imageBlock()`, а не `placeholder()`.

### `imageBlock(imagePathOrBuffer, width, height, opts)`

Вставляет изображение. Первый аргумент может быть путем к файлу или `Buffer`.

```js
imageBlock("/absolute/path/graph.png", 600, 420);
imageBlock(fs.readFileSync(graphPath), 600, 420);
imageBlock(fs.readFileSync(photoPath), 600, 420, { imageType: "jpg" });
```

Практическое правило: в отчетах надежнее передавать `Buffer`, особенно если путь собирается динамически.

```js
imageBlock(fs.readFileSync(plotPath), 600, 420);
```

По умолчанию используется `imageType: "png"`. Для других форматов передавайте `imageType: "jpg"`, `"gif"` или `"bmp"` в `opts`.

## DOCX: листинги кода

### `codeLine(text, opts)`

Создает одну строку кода без красной строки, моноширинным шрифтом.

```js
codeLine("int main() { return 0; }");
codeLine("console.log(value);", { size: 20 });
```

### `codeBlock(lines, opts)`

Создает массив строк кода. Обычно раскрывается через spread.

```js
const source = fs.readFileSync("program.cpp", "utf8").split("\n");

const children = [
  h1("Приложение А"),
  paragraph("Листинг программы."),
  ...codeBlock(source, { size: 20 }),
];
```

Если нужен особый размер только для одного листинга, передавайте `size` в `codeBlock()`. Глобальный `SIZE_CODE` менять не обязательно.

```js
...codeBlock(source, { size: 20 }) // 10 pt
```

## DOCX: математический API

Математические функции экспортируются отдельно из модуля, а не из экземпляра, потому что они не зависят от стиля.

```js
const {
  createDocxGost,
  mathFraction,
  mathSub,
  mathSup,
  mathSubSup,
  mathParen,
} = require("@codex/gost/docx");
```

### Основной принцип

Любая math-функция принимает строки, числа, массивы или другие math-компоненты.

```js
mathFraction("48", "13");
mathSub("x", "1");
mathSup("q", "2");
mathParen(["a + b"]);
```

### `toMathComponents(value)`

Внутренний нормализатор, но доступен публично. Преобразует строку, число, массив или math-компонент в массив компонентов `docx`.

Обычно напрямую не нужен.

### `mathExpr(...parts)`

Создает полноценный `docx.Math` из фрагментов.

```js
mathExpr("ν = ", mathFraction("48", "13"));
```

В большинстве отчетов удобнее использовать `formulaMath()`, который сам вызывает `mathExpr()`.

### `mathFraction(numerator, denominator)`

Дробь.

```js
formulaMath(["q* = ", mathFraction("10", "13")]);
```

### `mathSub(base, subScript)`

Нижний индекс.

```js
paragraphWithMath(["Стратегия ", mathSub("a", "1"), " активна."]);
```

### `mathSup(base, superScript)`

Верхний индекс.

```js
formulaMath([mathSup("q", "2"), " + 1"]);
```

### `mathSubSup(base, subScript, superScript)`

Нижний и верхний индекс одновременно.

```js
formulaMath([mathSubSup("x", "i", "*"), " = 0"]);
```

### `mathParen(...parts)`, `mathBracket(...parts)`, `mathBrace(...parts)`

Скобки: круглые, квадратные, фигурные.

```js
formulaMath(["f = ", mathParen("x - ", mathSub("x", "min"))]);
```

### `mathRoot(content, degree)`

Корень. Если `degree` не указан, получается квадратный корень.

```js
formulaMath(["G = ", mathRoot("a*b*c", "3")]);
```

### `mathSum(content, opts)`

Сумма с опциональными нижним и верхним пределами.

```js
formulaMath([
  mathSum(mathSub("x", "i"), { subScript: "i=1", superScript: "n" }),
]);
```

### `mathIntegral(content, opts)`

Интеграл с опциональными пределами.

```js
formulaMath([
  mathIntegral("f(x)dx", { subScript: "0", superScript: "1" }),
]);
```

### `formulaMath(content, opts)`

Отдельная центрированная формула.

```js
formulaMath(["Cmax = 23"]);
formulaMath(["ν = ", mathFraction("48", "13"), " ≈ 3.692"]);
formulaMath([mathSub("V", "a1"), "(q) = -4 + 10q"]);
```

### `formulaInline(label, mathContent, opts)`

Абзац вида `label + формула`. Использовать, когда нужен текстовый префикс перед математическим выражением.

```js
formulaInline("Получаем:", ["q* = ", mathFraction("10", "13")]);
```

### `paragraphWithMath(parts, opts)` vs `formulaMath(content, opts)`

Использовать `paragraphWithMath()`, если формула встроена в предложение:

```js
paragraphWithMath([
  "При ",
  mathSub("q", "*"),
  " активными являются строки ",
  mathSub("a", "1"),
  " и ",
  mathSub("a", "2"),
  ".",
]);
```

Использовать `formulaMath()`, если выражение должно стоять отдельной строкой:

```js
formulaMath(["-4 + 10q = 6 - 3q"]);
formulaMath(["13q = 10"]);
formulaMath(["q* = ", mathFraction("10", "13")]);
```

## DOCX: таблицы

### `makeTable(rows, opts)`

Создает таблицу `docx.Table`.

```js
makeTable([
  ["Критерий", "Значение"],
  ["MM", "E1"],
  ["BL", "E2"],
], { widths: [4000, 5355] });
```

Важные правила реализации:

- ширина таблицы задается только через `WidthType.DXA`;
- ширина ячеек задается только через `WidthType.DXA`;
- `WidthType.PERCENTAGE` не используется;
- первая строка по умолчанию считается заголовком;
- заголовочная строка по умолчанию центрируется и выделяется жирным;
- ширины колонок лучше задавать явно через `widths`.

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
| `borders` | объект границ docx |

### Типы ячеек

Ячейка может быть строкой, числом, math-компонентом или объектом:

```js
makeTable([
  ["Показатель", "Значение"],
  [{ text: "Итого", bold: true }, { text: 23, bold: true }],
]);
```

Объект ячейки:

| Поле | Смысл |
| --- | --- |
| `text` | значение ячейки: строка, число, массив фрагментов или math-компонент |
| `align` | выравнивание |
| `bold` | жирный |
| `italics` | курсив |
| `size` | размер |
| `color` | цвет |
| `font` | шрифт |
| `colspan` | объединение колонок |
| `rowspan` | объединение строк |
| `shading` | заливка в формате `docx` |

Пример с формулой в ячейке:

```js
makeTable([
  ["Формула", "Результат"],
  [[mathSub("V", "a1"), "(q)"], mathFraction("48", "13")],
]);
```

Правильный паттерн для таблиц в отчете:

```js
paragraph(`Результаты расчета приведены в таблице ${tables.peek()}.`);
tables.caption("результаты расчета");
makeTable(rows, { widths: [3000, 3000, 3355] });
```

## DOCX: титульный лист, секции и документ

### `makeTitlePage(opts)`

Создает первую секцию документа с титульным листом и нижним колонтитулом `Таганрог {year}`. Возвращаемый тип: `DocxSection`.

```js
makeTitlePage({
  workType: "ИНДИВИДУАЛЬНОЕ ЗАДАНИЕ",
  subtitle: ["по теории принятия решений", "вариант 2"],
  discipline: "по дисциплине: «Теория принятия решений»",
  group: "КТбо3-7",
  author: "Бораев Р. Х.",
  teacher: "Родзин С. И.",
  teacherTitle: "Доцент кафедры",
  year: 2026,
});
```

Поля:

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

В библиотеке нет hardcoded ФИО, групп и вариантов. Эти данные всегда передаются через `makeTitlePage()`.

### `makeContentSection(children)`

Создает секцию основного содержания с нумерацией страниц в нижнем колонтитуле. Возвращаемый тип: `DocxSection`.

```js
makeContentSection([
  h1("Введение"),
  paragraph("Текст введения."),
]);
```

### `makeDocument(sections, opts)`

Создает `docx.Document` из массива `DocxSection[]`.

```js
const doc = makeDocument([
  makeTitlePage(titleOpts),
  makeContentSection(children),
]);
```

По умолчанию:

- `settings.updateFields = true`;
- стили Heading 1/2/3 создаются автоматически;
- `TableOfContents` обновляется Word при открытии документа.

Можно передать свои стили `docx`:

```js
makeDocument(sections, { styles: customStyles });
```

### `saveDocument(doc, outputPath)`

Сохраняет документ на диск.

```js
await saveDocument(doc, path.join(__dirname, "..", "gen", "report.docx"));
```

Метод возвращает `Promise<void>`.

### `TableOfContents`

`TableOfContents` реэкспортируется из `docx`.

```js
const { TableOfContents } = require("@codex/gost/docx");

new TableOfContents("Содержание", {
  hyperlink: false,
  headingStyleRange: "1-3",
});
```

## DOCX: рекомендуемый шаблон отчета

```js
const g = createDocxGost();
const {
  h1, h2, paragraph, makeTable, makeTitlePage,
  makeContentSection, makeDocument, saveDocument,
  createCaptionCounter, pageBreak,
} = g;

const tables = createCaptionCounter("Таблица", { table: true });
const figures = createCaptionCounter("Рисунок");

const children = [
  h1("Содержание"),
  new TableOfContents("Содержание", { hyperlink: false, headingStyleRange: "1-3" }),
  pageBreak(),

  h1("Введение"),
  paragraph("Во введении описывается цель работы."),

  h1("1 Теоретическая часть"),
  h2("1.1 Исходные данные"),
  paragraph(`Исходные данные представлены в таблице ${tables.peek()}.`),
  tables.caption("исходные данные"),
  makeTable(rows),

  h1("Заключение"),
  paragraph("В заключении приводятся основные результаты."),
];

const doc = makeDocument([
  makeTitlePage(titleOpts),
  makeContentSection(children),
]);

await saveDocument(doc, outputPath);
```

## DOCX: частые ошибки

| Ошибка | Как правильно |
| --- | --- |
| Подпись таблицы через `caption()` | `createCaptionCounter("Таблица", { table: true })` или `tableCaption()` |
| Рисунок без ссылки в предыдущем абзаце | `paragraph(\`На рисунке ${figures.peek()} показан ...\`)` |
| Таблица без ссылки в предыдущем абзаце | `paragraph(\`... приведено в таблице ${tables.peek()}.\`)` |
| Математика plain text внутри `paragraph()` | `formulaMath()` или `paragraphWithMath()` |
| `imageBlock(pathString)` с ненадежным путем | `imageBlock(fs.readFileSync(path), w, h)` |
| Листинг кода в середине раздела вместо приложения | сослаться на приложение и вставить `codeBlock()` в приложении |
| Глобальные счетчики | локальные `const figures = createCaptionCounter(...)` внутри сборки документа |

## Excel: минимальный пример

```js
const path = require("path");
const ExcelJS = require("exceljs");
const { createExcelGost } = require("@codex/gost/excel");

const x = createExcelGost({
  fontSize: 14,
  headerFill: "FFF2CC",
});

const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet("Матрица");

x.applyBaseSheetSetup(sheet, [20, 18, 18, 18]);
x.mergeTitle(sheet, "A1:D1", "ИСХОДНАЯ МАТРИЦА");

x.setValue(sheet, "A3", "Стратегия", x.HEADER_STYLE);
x.setValue(sheet, "B3", "b1", x.HEADER_STYLE);
x.setValue(sheet, "C3", "b2", x.HEADER_STYLE);
x.setValue(sheet, "D3", "b3", x.HEADER_STYLE);

x.setValue(sheet, "A4", "a1", { bold: true, border: true, fill: x.style.HEADER_FILL });
x.setValue(sheet, "B4", 3, { border: true });
x.setValue(sheet, "C4", 5, { border: true });
x.setValue(sheet, "D4", 7, { border: true });

x.formatTableRange(sheet, 3, 4, 1, 4, 3);
x.autoFitSheet(sheet);

await workbook.xlsx.writeFile(path.join(__dirname, "..", "gen", "report.xlsx"));
```

## Excel: архитектура

`createExcelGost(opts)` возвращает объект с методами для оформления workbook/sheet. Библиотека не создает workbook сама, потому что это ответственность `exceljs`.

```js
const ExcelJS = require("exceljs");
const { createExcelGost } = require("@codex/gost/excel");

const workbook = new ExcelJS.Workbook();
const x = createExcelGost();
```

## Excel: настройка стиля

### `createExcelStyle(opts)`

Создает полный объект стиля Excel. Принимает поля `ExcelStyleConfig` в верхнем регистре.

```js
const style = createExcelStyle({
  FONT_SIZE: 14,
  HEADER_FILL: "D9EAF7",
  RESULT_FILL: "D9EAD3",
});
```

Поля стиля:

| Поле | Значение по умолчанию | Смысл |
| --- | --- | --- |
| `HEADER_FILL` | `FFF2CC` | заливка заголовков и меток |
| `RESULT_FILL` | `D9EAD3` | заливка итоговых ячеек |
| `ALT_ROW_FILL` | `FFF9F0` | заливка четных строк при zebra striping |
| `REGRET_FILL` | `FCE4D6` | заливка для матриц сожалений |
| `FONT_NAME` | `Times New Roman` | основной шрифт |
| `FONT_SIZE` | `14` | основной размер |
| `TAB_COLOR` | `FFF2CC` | цвет вкладки листа |
| `DEFAULT_ROW_HEIGHT` | `20` | высота строки по умолчанию |
| `PAGE_ORIENTATION` | `landscape` | ориентация страницы |
| `FIT_TO_PAGE` | `true` | включить подгонку под страницу |
| `FIT_TO_WIDTH` | `1` | подгонка по ширине |
| `FIT_TO_HEIGHT` | `0` | подгонка по высоте, 0 - без ограничения |
| `MAX_COLUMN_WIDTH` | `80` | максимум ширины столбца при autofit |
| `MAX_ROW_HEIGHT` | `120` | максимум высоты строки при autofit |
| `ROW_HEIGHT_FONT_MULTIPLIER` | `1.9` | множитель высоты строки |
| `HEADER_TITLE_SIZE` | `14` | размер заголовков листов |
| `BORDER_STYLE` | `thin` | стиль границ по умолчанию |

### Alias-опции `createExcelGost(opts)`

`createExcelGost()` принимает те же поля в верхнем регистре и удобные alias-поля. Если переданы оба варианта, поле в верхнем регистре имеет приоритет.

```js
const x = createExcelGost({
  fontSize: 14,
  fontName: "Arial",
  headerFill: "D9EAF7",
  resultFill: "D9EAD3",
  tabColor: "D9EAF7",
});

const y = createExcelGost({
  fontSize: 12,
  FONT_SIZE: 14, // победит FONT_SIZE
});
```

| Alias | Полное поле |
| --- | --- |
| `fontSize` | `FONT_SIZE` |
| `fontName` | `FONT_NAME` |
| `headerFill` | `HEADER_FILL` |
| `resultFill` | `RESULT_FILL` |
| `tabColor` | `TAB_COLOR` |

### `HEADER_STYLE` и `RESULT_STYLE`

Экземпляр `createExcelGost()` содержит готовые стили:

```js
x.HEADER_STYLE // { bold: true, border: true, fill: HEADER_FILL, align: "center" }
x.RESULT_STYLE // { bold: true, border: "medium", fill: RESULT_FILL, align: "center" }
```

Для левых меток используйте тот же стиль, но с `align: "left"`:

```js
const LABEL_STYLE = {
  bold: true,
  border: true,
  fill: x.style.HEADER_FILL,
  align: "left",
};
```

## Excel: стили ячеек

### `styleCell(cell, options)`

Применяет стиль к существующей ячейке.

```js
x.styleCell(sheet.getCell("A1"), {
  bold: true,
  fill: x.style.HEADER_FILL,
  border: true,
  align: "center",
});
```

Опции:

| Поле | Смысл |
| --- | --- |
| `fontName` | шрифт |
| `size` | размер |
| `bold` | жирный |
| `color` | цвет ARGB/RGB строкой |
| `align` | `left`, `center`, `right`, `justify` и т.п. |
| `wrapText` | перенос текста, по умолчанию включен |
| `fill` | цвет заливки |
| `border` | `true` или стиль границы, например `medium` |
| `numFmt` | числовой формат Excel |

### `setValue(sheet, address, value, options)`

Записывает значение и сразу применяет стиль.

```js
x.setValue(sheet, "A1", "Показатель", x.HEADER_STYLE);
x.setValue(sheet, "B1", 12.345, { border: true, numFmt: "0.000" });
```

### `setFormula(sheet, address, formula, options)`

Записывает формулу без начального `=`.

```js
x.setFormula(sheet, "B10", "SUM(B4:B9)", {
  border: true,
  numFmt: "0.0000",
});
```

## Excel: настройка листа

### `applyBaseSheetSetup(sheet, widths)`

Применяет базовые параметры листа:

- высота строк по умолчанию;
- ориентация страницы;
- fit-to-page;
- ширины колонок;
- запоминает базовые ширины для `autoFitSheet()`.

```js
x.applyBaseSheetSetup(sheet, [18, 14, 14, 14, 20]);
```

### `mergeTitle(sheet, range, text)`

Объединяет диапазон и пишет крупный заголовок.

```js
x.mergeTitle(sheet, "A1:F1", "РАСЧЕТ КРИТЕРИЕВ");
```

### `setFreeze(sheet, ySplit, activeCell)`

Замораживает верхние `ySplit` строк через `sheet.views`. `activeCell` задает активную ячейку после заморозки.

```js
x.setFreeze(sheet, 3, "A4");
```

### `applyZebraStriping(sheet, startRow, endRow, startCol, endCol, headerRows)`

Применяет мягкую заливку четных строк. Строки из `headerRows` пропускаются.

```js
x.applyZebraStriping(sheet, 3, 20, 1, 8, new Set([3]));
```

### `autoFitSheet(sheet)`

Автоматически подбирает ширину столбцов и высоту строк по тексту, учитывая ограничения стиля.

```js
x.autoFitSheet(sheet);
```

Использовать в конце формирования листа.

### `formatTableRange(sheet, startRow, endRow, startCol, endCol, headerRow)`

Оформляет прямоугольный диапазон как таблицу: границы всем ячейкам, заголовочной строке жирный шрифт и заливка.

```js
x.formatTableRange(sheet, 3, 12, 1, 6, 3);
```

### `getCellText(cell)`

Возвращает текстовое представление значения ячейки. Формулы возвращают пустую строку, потому что у них нет рассчитанного результата внутри `exceljs`.

```js
const text = x.getCellText(sheet.getCell("A1"));
```

## Excel: формульные помощники

Формульные помощники доступны как методы экземпляра `createExcelGost()`. Они не экспортируются напрямую из модуля.

### `productFormula(addresses, exponent)`

Создает формулу геометрического среднего.

```js
x.productFormula(["B4", "C4", "D4"], 3);
// (B4*C4*D4)^(1/3)
```

### `absoluteRef(address)`

Преобразует адрес в абсолютную ссылку.

```js
x.absoluteRef("B12"); // $B$12
```

Если адрес не похож на `A1`, возвращается как есть.

### `weightedSumFormula(valueAddresses, weightAddresses)`

Создает сумму произведений.

```js
x.weightedSumFormula(["B2", "C2", "D2"], ["$B$1", "$C$1", "$D$1"]);
// B2*$B$1+C2*$C$1+D2*$D$1
```

### `listFormulaRange(sheetName, addresses)`

Создает список ссылок на ячейки с указанием листа.

```js
x.listFormulaRange("МАИ", ["B2", "B3"]);
// 'МАИ'!B2,'МАИ'!B3
```

## Excel: блок МАИ

### `addAhpMatrixBlock(sheet, startRow, title, labels, matrix, randomIndex)`

Строит блок матрицы парных сравнений МАИ:

- заголовок блока;
- подписи строк и столбцов;
- матрица сравнений;
- геометрические средние `G_i`;
- веса `w_i`;
- сумма весов;
- суммы столбцов;
- `λ_max`;
- индекс согласованности `ИС`;
- случайный индекс `ИС_сл`;
- отношение согласованности `ОС`;
- `ОС%`.

`labels` должен быть непустым массивом, а `matrix` - квадратной матрицей того же размера. При несоответствии метод бросает `RangeError` до записи данных в лист. Колонки рассчитываются через внутренний `colLetter()`, поэтому блок работает и после столбца `Z`.

```js
const result = x.addAhpMatrixBlock(
  sheet,
  5,
  "Матрица парных сравнений критериев",
  ["К1", "К2", "К3"],
  [
    [1, 3, 5],
    [1 / 3, 1, 2],
    [1 / 5, 1 / 2, 1],
  ],
  0.58,
);
```

Возвращает объект:

```ts
{
  weightColumn: string;
  weightStartRow: number;
  weightEndRow: number;
  osCell: string;
  nextRow: number;
}
```

Пример использования результата:

```js
const block = x.addAhpMatrixBlock(sheet, 5, title, labels, matrix, 1.24);
x.setFormula(sheet, "B40", `${block.osCell}*100`, { numFmt: "0.00" });
```

## Excel: содержание workbook

### `addContentsSheet(workbook, sheetNames)`

Создает лист `Содержание` со ссылками на остальные листы.

```js
x.addContentsSheet(workbook, [
  "Матрица",
  "Критерии",
  "МАИ",
]);
```

Обычно вызывать после создания всех расчетных листов или перед ними, если список имен уже известен.

## Excel: рекомендуемый шаблон workbook

```js
const ExcelJS = require("exceljs");
const { createExcelGost } = require("@codex/gost/excel");

async function buildWorkbook(outputPath) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "report-generator";
  workbook.created = new Date();

  const x = createExcelGost({ fontSize: 14 });
  const sheet = workbook.addWorksheet("Расчет", {
    properties: { tabColor: { argb: x.style.TAB_COLOR } },
  });

  x.applyBaseSheetSetup(sheet, [24, 18, 18, 18]);
  x.mergeTitle(sheet, "A1:D1", "РАСЧЕТ");

  const LABEL_STYLE = {
    bold: true,
    border: true,
    fill: x.style.HEADER_FILL,
    align: "left",
  };

  x.setValue(sheet, "A3", "N=", LABEL_STYLE);
  x.setValue(sheet, "B3", 7, { border: true });

  x.setValue(sheet, "A5", "Показатель", x.HEADER_STYLE);
  x.setValue(sheet, "B5", "Значение", x.HEADER_STYLE);
  x.setValue(sheet, "A6", "Cmax", { border: true, align: "left" });
  x.setFormula(sheet, "B6", "SUM(B3:B3)", { border: true, numFmt: "0.000" });

  x.formatTableRange(sheet, 5, 6, 1, 2, 5);
  x.autoFitSheet(sheet);

  await workbook.xlsx.writeFile(outputPath);
}
```

## Excel: частые ошибки

| Ошибка | Как правильно |
| --- | --- |
| Метки слева без заливки и границ | `{ bold: true, border: true, fill: x.style.HEADER_FILL, align: "left" }` |
| Формула с начальным `=` | `setFormula(sheet, "A1", "SUM(B1:B2)")` |
| Забыли `autoFitSheet()` | вызвать в конце каждого листа |
| Нечитаемый мелкий шрифт | `createExcelGost({ fontSize: 14 })` |
| Workbook создается библиотекой | workbook создает пользователь через `new ExcelJS.Workbook()` |
| Нужен итоговый стиль | использовать `x.RESULT_STYLE` |

## Импорт из TypeScript

```ts
import {
  createDocxGost,
  TableOfContents,
  mathFraction,
  mathSub,
} from "@codex/gost/docx";

import { createExcelGost } from "@codex/gost/excel";
```

Для CommonJS используйте те же package subpaths через `require()`.

## Импорт из CommonJS

```js
const {
  createDocxGost,
  TableOfContents,
  mathFraction,
} = require("@codex/gost/docx");

const { createExcelGost } = require("@codex/gost/excel");
```

## Legacy API

Файлы `docx-gost.js` и `excel-gost.js` в корне `lib` экспортируют функции напрямую, без фабрики:

```js
const {
  paragraph,
  h1,
  makeTable,
  makeDocument,
} = require("../../lib/docx-gost");
```

Для нового кода лучше использовать factory API:

```js
const { createDocxGost } = require("@codex/gost/docx");
const g = createDocxGost();
```

Причина: фабрика позволяет менять константы через опции и не создает скрытого общего состояния между документами.

## Правила для генератора отчетов

Если эту документацию использует нейронка для написания нового отчета, придерживаться следующих правил:

1. Все DOCX-элементы класть в массив `children`.
2. Заголовки делать через `h1`, `h2`, `h3`, а не жирными абзацами.
3. Таблицу всегда предварять абзацем-ссылкой и `tables.caption()`.
4. Для таблиц создавать счетчик только так: `createCaptionCounter("Таблица", { table: true })`.
5. Рисунок всегда предварять абзацем-ссылкой, затем `imageBlock()` или `placeholder()`, затем `figures.caption()`.
6. Формулы и математические выражения не писать plain text в `paragraph()`, если есть индексы, дроби, знаки сравнения, стрелки, равенства или приближения.
7. Отдельные формулы оформлять через `formulaMath()`.
8. Формулы внутри предложения оформлять через `paragraphWithMath()`.
9. Исходный код больших программ лучше помещать в приложение через `codeBlock()`.
10. Для Excel все ячейки-метки слева оформлять стилем `LABEL_STYLE`.
11. Для Excel все таблицы оформлять через `formatTableRange()`.
12. В конце каждого Excel-листа вызывать `autoFitSheet()`.
13. Не хранить ФИО, группы, номера вариантов и другие данные внутри библиотеки. Передавать их в скрипт сборки.
14. Не добавлять side effects при импорте: генерация файлов должна запускаться только явным вызовом функции или main-скрипта.

## PlantUML: диаграммы в коде

Модуль позволяет писать PlantUML прямо в скрипте отчёта — никаких внешних `.puml` файлов. Диаграмма рендерится на лету и вставляется в DOCX.

```js
const { createPlantUmlGost } = require("@magomed-cmd/gost/plantuml");

const p = createPlantUmlGost(); // dpi: 150 по умолчанию
```

Требует установленной Java (для node-plantuml).

### Builders

Все builders поддерживают `.add(line)` как escape hatch для любой строки PlantUML.

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

#### `p.stateDiagram()`

Метод `transition` принимает опциональное направление `"down"`, `"up"`, `"left"`, `"right"` — помогает избежать пересечений стрелок.

```js
p.stateDiagram()
  .initial("ON_LINE")
  .state("ON_LINE", "На линии")
  .state("BROKEN",  "Сломан")
  .state("IN_DEPOT","В парке")
  .state("RESERVE", "Резерв")
  .transition("ON_LINE", "BROKEN",   "поломка",         "down")
  .transition("BROKEN",  "IN_DEPOT", "отремонтирован")
  .transition("IN_DEPOT","ON_LINE",  "выход на линию")
  .transition("ON_LINE", "RESERVE",  "избыток")
  .transition("RESERVE", "ON_LINE",  "нехватка")
  .final("IN_DEPOT")
  .build()
```

### `p.render(puml, opts?)`

Рендерит puml-строку, возвращает PNG Buffer. Базовые skinparam (Times New Roman 12pt, dpi 150) добавляются автоматически.

```js
const png = await p.render(p.classDiagram().class("Foo").build());
```

Опции: `format` (`"png"` или `"svg"`), `dpi`, `skinparams`.

### `autoImageSize(buffer, maxWidth?)`

Читает реальные размеры PNG из буфера и масштабирует пропорционально. Никакого хардкода размеров.

```js
const { autoImageSize } = require("@magomed-cmd/gost/plantuml");

const png = await p.render(puml);
const { width, height } = autoImageSize(png);       // maxWidth = 624 (ГОСТ A4)
const { width, height } = autoImageSize(png, 400);  // своя ширина

imageBlock(png, width, height);
figures.caption("диаграмма классов");
```

Максимальная ширина контента ГОСТ A4 = **624px** (9355 DXA при 150 DPI).

### `wrapWithSkin(source, dpi?, opts?)`

Добавляет базовые skinparam к puml-строке. Вызывается внутри `render()` автоматически, но доступен публично для ручного использования.

```js
const { wrapWithSkin } = require("@magomed-cmd/gost/plantuml");

const styled = wrapWithSkin(rawPuml, 300, { skinparams: { backgroundColor: "white" } });
```

Базовые skinparam: Times New Roman 12pt, Shadowing false, nodesep 60, ranksep 50.

---

## DOCX: `diagramBlock`

Комбо-метод на инстансе `g`: рендерит PlantUML строку → авторазмер → `imageBlock` + подпись. Спредится прямо в `children`.

```js
const { createDocxGost } = require("@magomed-cmd/gost/docx");
const { createPlantUmlGost } = require("@magomed-cmd/gost/plantuml");

const g = createDocxGost();
const p = createPlantUmlGost();
const figures = g.createCaptionCounter("Рисунок");

const children = [
  paragraph(`На рисунке ${figures.peek()} представлена диаграмма последовательности.`),
  ...await g.diagramBlock(
    p.sequenceDiagram()
      .actor("Пользователь", "U")
      .participant("Dispatcher", "D")
      .arrow("U", "D", "triggerBreakdown()")
      .build(),
    figures,
    "диаграмма последовательности обработки поломки"
  ),
];
```

Опции `DiagramBlockOpts`:

| Поле | По умолчанию | Смысл |
| --- | --- | --- |
| `dpi` | 150 | качество растеризации (300 для финального отчёта на печать) |
| `maxWidth` | 624 | максимальная ширина px |
| `skinparams` | — | переопределение skinparam |

Если диаграмма выше 900px — в консоль выводится предупреждение.

---

## DOCX: `lint` — проверка документа

`g.lint(children)` анализирует массив **до** сборки документа и выводит предупреждения в консоль. Возвращает массив `LintIssue[]`.

```js
g.lint(children); // выводит предупреждения
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
| `empty-page-start` | warn | `pageBreak` первым элементом — пустая страница в начале |
| `empty-page-end` | warn | `pageBreak` последним элементом — пустая страница в конце |
| `consecutive-page-breaks` | warn | два `pageBreak` подряд |
| `page-break-blank-page-break` | warn | `pageBreak` → `blank(и)` → `pageBreak` |
| `heading-before-page-break` | warn | `h1/h2/h3` прямо перед `pageBreak` — заголовок повиснет внизу страницы |
| `image-caption-split` | warn | `image` → `pageBreak` → `caption` — рисунок и подпись на разных страницах |
| `table-caption-split` | warn | `tableCaption` → `pageBreak` — подпись таблицы отделена от таблицы |
| `image-without-caption` | warn | `image` без `caption` следом (нарушение ГОСТ) |
| `placeholder-left` | warn | `placeholder` в финальном документе |
| `blank-before-page-break` | info | `blank` прямо перед `pageBreak` — лишний элемент |
| `too-many-blanks` | info | 3+ `blank` подряд |

Для программного использования (например в CI):

```js
const { lintChildren, printLintResults } = require("@magomed-cmd/gost/docx");

const issues = lintChildren(children);
printLintResults(issues);
if (issues.filter(i => i.level === "warn").length > 0) process.exit(1);
```

Замечание по `imageBlock`: по умолчанию `keepNext: true`, что предотвращает **естественный** разрыв страницы между рисунком и подписью. `image-caption-split` ловит только **явные** `pageBreak()`.

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

## Полный список экспортов DOCX

Из `@codex/gost/docx`:

```js
{
  DEFAULT_DOCX_STYLE,
  createDocxStyle,
  createDocxGost,
  lintChildren,
  printLintResults,
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
  TableOfContents,
}
```

Методы экземпляра `createDocxGost()`:

```js
{
  style,
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
  formulaMath,
  formulaInline,
  makeTable,
  makeTitlePage,
  makeContentSection,
  makeDocument,
  saveDocument,
  diagramBlock,
  lint,
}
```

## Полный список экспортов Excel

Из `@codex/gost/excel`:

```js
{
  createExcelStyle,
  createExcelGost,
}
```

Методы экземпляра `createExcelGost()`:

```js
{
  style,
  HEADER_STYLE,
  RESULT_STYLE,
  styleCell,
  setValue,
  setFormula,
  applyBaseSheetSetup,
  mergeTitle,
  setFreeze,
  applyZebraStriping,
  getCellText,
  autoFitSheet,
  formatTableRange,
  addAhpMatrixBlock,
  addContentsSheet,
  productFormula,
  absoluteRef,
  weightedSumFormula,
  listFormulaRange,
}
```

## Полный список экспортов PlantUML

Из `@magomed-cmd/gost/plantuml`:

```js
{
  createPlantUmlGost,
  autoImageSize,
  renderDiagram,
  wrapWithSkin,
  UseCaseBuilder,
  ClassDiagramBuilder,
  SequenceBuilder,
  StateBuilder,
}
```

Методы экземпляра `createPlantUmlGost()`:

```js
{
  useCaseDiagram,
  classDiagram,
  sequenceDiagram,
  stateDiagram,
  render,
  autoImageSize,
}
```

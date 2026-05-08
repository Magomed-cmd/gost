# ПЛАН: библиотека `docx-gost.js`

## Цель

Создать локальную библиотеку `lib/docx-gost.js` для генерации `.docx` отчётов  
по ГОСТ (ЮФУ, ИКТИБ) на основе `docx` npm-пакета.

Принцип: каждый скрипт отчёта делает `require('../lib/docx-gost')` и получает  
все хелперы — больше не нужно копипастить 500 строк в каждый файл.

---

## Файловая структура

```
CODEX/
├── lib/
│   ├── docx-gost.js      ← сама библиотека (создать)
│   └── README.md         ← документация с примерами (создать)
├── variant2/
│   └── scripts/
│       └── build_report.js   ← будущий скрипт, использует lib
└── variant4_archive/
    └── scripts/report/
        └── build_report.js   ← старый скрипт (только читать, не трогать)
```

---

## Зависимости

```json
{ "docx": ">=8.0.0" }
```

`docx-gost.js` не устанавливает зависимости сам — `docx` должен быть установлен  
в проекте, который использует библиотеку (`npm install docx`).

---

## Структура `docx-gost.js`

Файл разбит на секции. Каждая секция экспортирует свои функции через  
`module.exports` в конце файла (единый объект).

---

## Секция 1 — Константы

```js
const FONT = "Times New Roman";
const SIZE = 28;          // 14pt — основной текст
const SIZE_TITLE = 32;    // 16pt — "ОТЧЁТ" на титульнике
const SIZE_HEADER = 24;   // 12pt — шапка титульника
const SIZE_CODE = 24;     // 12pt — листинги (Courier New)
const INDENT = 709;       // 1.25 см — красная строка
const LINE = 360;         // 1.5 интервал
const CONTENT_WIDTH = 9355; // ширина текста в DXA (11906 - 1701 - 850)

const PAGE = {
  size: { width: 11906, height: 16838 },   // A4
  margin: { top: 1134, bottom: 1134, left: 1701, right: 850 }
  // top=2cm, right=1.5cm, bottom=2cm, left=3cm
};

const BORDER_SINGLE = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
const BORDERS_ALL = {
  top: BORDER_SINGLE, bottom: BORDER_SINGLE,
  left: BORDER_SINGLE, right: BORDER_SINGLE
};
```

Экспортировать: `FONT, SIZE, SIZE_TITLE, SIZE_HEADER, SIZE_CODE, INDENT, LINE, CONTENT_WIDTH, PAGE, BORDER_SINGLE, BORDERS_ALL`

---

## Секция 2 — Текстовые примитивы

### `run(text, opts = {})`

Возвращает `TextRun`. Параметры `opts`:
- `size` — размер (default: `SIZE`)
- `bold` — жирный (default: false)
- `italics` — курсив (default: false)
- `allCaps` — все заглавные (default: false)
- `underline` — подчёркивание (default: false)
- `color` — цвет hex без # (default: undefined)
- `font` — шрифт (default: `FONT`)

### `paragraph(content, opts = {})`

Возвращает `Paragraph`. `content` — строка или массив `{text, bold, italics, size, ...}`.

Параметры `opts`:
- `align` — выравнивание (default: `AlignmentType.JUSTIFIED`)
- `noIndent` — без красной строки (default: false)
- `indent` — размер отступа в twips (default: `INDENT`)
- `before`, `after` — отступы в twips (default: 0)
- `line` — межстрочный (default: `LINE`)
- `keepNext` — не отрывать от следующего (default: false)
- `keepLines` — не разбивать абзац (default: **true**) ← всегда включён
- `heading` — `HeadingLevel.HEADING_X` (default: undefined)
- `bold`, `italics`, `size` — применяются ко всем TextRun, если content — строка

### `centered(content, opts = {})`

Shortcut: `paragraph(content, { ...opts, align: CENTER, noIndent: true })`

---

## Секция 3 — Заголовки

### `h1(text)`

- `HeadingLevel.HEADING_1`
- Выравнивание: CENTER
- Текст: `text.toUpperCase()`, bold, color `"000000"`
- `before: 240, after: 120, keepNext: true, noIndent: true`

### `h2(text)`

- `HeadingLevel.HEADING_2`
- Выравнивание: JUSTIFIED (с красной строкой)
- bold, color `"000000"`
- `before: 200, after: 80, keepNext: true`

### `h3(text)`

- `HeadingLevel.HEADING_3`
- bold + italics, color `"000000"`
- `before: 140, after: 60, keepNext: true`

---

## Секция 4 — Структурные элементы

### `blank()`

Пустая строка: `Paragraph` с пустым `TextRun`, `line: LINE`.

### `pageBreak()`

`Paragraph` с `PageBreak` внутри.

### `caption(text)`

Подпись к рисунку/таблице:
- CENTER, `noIndent: true`
- `before: 60, after: 60`
- Размер `SIZE`

### `placeholder(label)`

Серая рамка-заглушка для рисунка (пока нет скриншота):
- CENTER, `noIndent: true`
- `before: 120, after: 60`
- Текст: italics, color `"888888"`
- Рамка: `BorderStyle.SINGLE, size: 4, color: "AAAAAA"` со всех сторон

### `imageBlock(imagePath, width, height, opts = {})`

Параграф с изображением по центру:
- Читает файл через `fs.readFileSync(imagePath)`
- `ImageRun` с `transformation: { width, height }`, `type: "png"`
- `before: 120, after: 60`

---

## Секция 5 — Листинг кода

### `codeLine(text)`

Одна строка кода:
- Courier New, `SIZE_CODE` (24 = 12pt)
- `noIndent: true, line: LINE`

### `codeBlock(lines)`

Принимает массив строк `string[]`, возвращает массив `Paragraph[]` через `.map(codeLine)`.

Пример использования:
```js
...codeBlock(["const x = 1;", "console.log(x);"])
```

---

## Секция 6 — Math

Все функции являются тонкими обёртками над `docx` Math-классами.  
Ключевая вспомогательная функция — `toMathComponents(value)`:
- если строка/число → `[new MathRun(String(value))]`
- если массив → рекурсивно flatten
- если объект docx → `[value]`

### Функции:

| Имя | Возвращает | Описание |
|-----|-----------|----------|
| `mathExpr(...parts)` | `Math` | Inline math-объект |
| `mathFraction(num, den)` | `MathFraction` | Дробь |
| `mathSub(base, sub)` | `MathSubScript` | Нижний индекс |
| `mathSup(base, sup)` | `MathSuperScript` | Верхний индекс |
| `mathSubSup(base, sub, sup)` | `MathSubSuperScript` | Оба индекса |
| `mathParen(...parts)` | `MathRoundBrackets` | Круглые скобки |
| `mathBracket(...parts)` | `MathSquareBrackets` | Квадратные скобки |
| `mathBrace(...parts)` | `MathCurlyBrackets` | Фигурные скобки |
| `mathRoot(content, degree?)` | `MathRadical` | Корень (степень опциональна) |
| `mathSum(content, opts?)` | `MathSum` | Сумма; `opts.subScript`, `opts.superScript` |
| `mathIntegral(content, opts?)` | `MathIntegral` | Интеграл; `opts.subScript`, `opts.superScript` |

### `formulaMath(content, opts = {})`

Отдельный центрированный абзац с формулой. `content` — массив частей для `mathExpr`.  
`before: 80, after: 80`, `noIndent: true`, `align: CENTER`.

### `formulaInline(label, mathContent, opts = {})`

Абзац: сначала текстовый `label`, потом `mathExpr(mathContent)` в одной строке.

### `paragraphWithMath(parts, opts = {})`

Абзац, где `parts` — смешанный массив строк и Math-объектов.  
Строки → `TextRun`, Math-объекты → вставляются напрямую.

---

## Секция 7 — Таблицы

### `makeTable(rows, opts = {})`

Возвращает `Table`.

`rows` — массив массивов ячеек: `[[cell, cell, ...], ...]`.  
Каждая ячейка — строка или объект `{ text, bold, align, colspan, rowspan, shading }`.

Параметры `opts`:
- `widths` — массив ширин столбцов в DXA; если не указан — равномерное деление `CONTENT_WIDTH`
- `fontSize` — размер шрифта в ячейках (default: `SIZE`)
- `headerRow` — индекс строки-заголовка (default: 0); она получает `bold: true`
- `borders` — объект бордеров (default: `BORDERS_ALL`)
- `align` — выравнивание текста в ячейках (default: JUSTIFIED)
- `centerHeader` — CENTER для строки-заголовка (default: true)

---

## Секция 8 — Структура документа

### `makeFooterCity(year)`

Footer с текстом `"Таганрог {year}"` по центру, Times New Roman SIZE.

### `makeFooterPageNum()`

Footer с `PageNumber.CURRENT` по центру.

### `makeTitlePage(opts)`

Возвращает объект секции (для `sections[]`) — титульная страница без нумерации.

Параметры `opts`:
```js
{
  workType: "ОТЧЁТ",              // "ОТЧЁТ" / "ИНДИВИДУАЛЬНОЕ ЗАДАНИЕ" / etc.
  subtitle: "о выполнении ...",   // строка или массив строк (многострочный)
  discipline: "по дисциплине: «...»",
  author: "Валиев М. М.",
  group: "КТбо3-7",
  teacher: "Иванов И. И.",
  teacherTitle: "Доцент кафедры", // default: "Доцент кафедры"
  year: 2026,
}
```

Структура раздела — реализовать ТОЧНО в таком порядке:

**1. Шапка университета** — 6 строк, SIZE_HEADER=24, CENTER, `noIndent: true`:
```
"Министерство науки и высшего образования Российской Федерации"
"Федеральное государственное автономное образовательное"
"учреждение высшего образования"
"«ЮЖНЫЙ ФЕДЕРАЛЬНЫЙ УНИВЕРСИТЕТ»"  ← bold
"Инженерно-технологическая академия"
"Институт компьютерных технологий и информационной безопасности"
```

**2.** 3 пустых строки (`blank()`)

**3. Тип работы** — `opts.workType`, SIZE_TITLE=32, bold, CENTER, `noIndent: true`

**4. Подзаголовок** — `opts.subtitle`: если строка — один абзац; если массив строк — по абзацу на каждую. SIZE, CENTER, `noIndent: true`

**5. Дисциплина** — `opts.discipline`, SIZE, CENTER, `noIndent: true`, `after: 120`

**6.** 9 пустых строк (`blank()`)

**7. Блок подписей** — два таб-стопа: `LEFT на 4500`, `RIGHT на 9355`:

```js
// Строка "Выполнил" — без таб-стопов, indent: { left: 709 }
new Paragraph({
  spacing: { line: LINE },
  indent: { left: 709 },
  children: [run("Выполнил")]
}),
// Строка с подписью студента
new Paragraph({
  spacing: { line: LINE },
  tabStops: [
    { type: TabStopType.LEFT, position: 4500 },
    { type: TabStopType.RIGHT, position: 9355 },
  ],
  children: [
    run(`студент группы ${opts.group}`),
    new Tab(),
    run("_______________"),
    new Tab(),
    run(opts.author),
  ]
}),
blank(),
// Строка "Принял" — без таб-стопов, indent: { left: 709 }
new Paragraph({
  spacing: { line: LINE },
  indent: { left: 709 },
  children: [run("Принял")]
}),
// Строка с подписью преподавателя
new Paragraph({
  spacing: { line: LINE, after: 560 },
  tabStops: [
    { type: TabStopType.LEFT, position: 4500 },
    { type: TabStopType.RIGHT, position: 9355 },
  ],
  children: [
    run(opts.teacherTitle ?? "Доцент кафедры"),
    new Tab(),
    run("_______________"),
    new Tab(),
    run(opts.teacher),
  ]
}),
```

**8.** Footer: `makeFooterCity(opts.year)`

### `makeContentSection(children, year?)`

Возвращает объект секции с нумерацией страниц.

`children` — массив `Paragraph | Table | ...`

Внутри уже задаётся:
- `properties: { page: PAGE }`
- Footer: `makeFooterPageNum()`
- `TableOfContents` НЕ включён — его добавляет пользователь сам в `children`

### `makeDocument(sections, opts = {})`

Возвращает готовый `Document`.

`opts.styles` — опционально переопределить стили (если не задано, применяются стандартные ГОСТ-стили: TNR 14pt, Heading1/Heading2 переопределены).

Стандартные стили включаются автоматически:
```js
styles: {
  default: { document: { run: { font: FONT, size: SIZE } } },
  paragraphStyles: [
    { id: "Heading1", ..., run: { size: 28, bold: true, color: "000000" }, paragraph: { spacing: { before: 240, after: 120 } } },
    { id: "Heading2", ..., run: { size: 28, bold: true, color: "000000" }, paragraph: { spacing: { before: 200, after: 80 } } },
    { id: "Heading3", ..., run: { size: 28, bold: true, color: "000000" }, paragraph: { spacing: { before: 140, after: 60 } } },
  ]
}
```

### `saveDocument(doc, outputPath)`

`Packer.toBuffer(doc).then(buffer => fs.writeFileSync(outputPath, buffer))` — промис.

---

## Экспорт

```js
module.exports = {
  // Константы
  FONT, SIZE, SIZE_TITLE, SIZE_HEADER, SIZE_CODE,
  INDENT, LINE, CONTENT_WIDTH, PAGE,
  BORDER_SINGLE, BORDERS_ALL,

  // Примитивы
  run, paragraph, centered, paragraphWithMath,

  // Заголовки
  h1, h2, h3,

  // Структурные
  blank, pageBreak, caption, placeholder, imageBlock,

  // Листинг
  codeLine, codeBlock,

  // Math
  toMathComponents, mathExpr, mathFraction,
  mathSub, mathSup, mathSubSup,
  mathParen, mathBracket, mathBrace,
  mathRoot, mathSum, mathIntegral,
  formulaMath, formulaInline,

  // Таблицы
  makeTable,

  // Документ
  makeFooterCity, makeFooterPageNum,
  makeTitlePage, makeContentSection,
  makeDocument, saveDocument,
};
```

---

## README.md — минимальный пример использования

```js
const {
  h1, h2, blank, body, centered, paragraph,
  codeLine, codeBlock, caption, placeholder, imageBlock,
  formulaMath, mathFraction, mathSub,
  makeTable, makeTitlePage, makeContentSection,
  makeDocument, saveDocument,
  pageBreak,
  TableOfContents, // реэкспортируется из docx
} = require('../lib/docx-gost');

const children = [
  h1("Содержание"),
  new TableOfContents("Содержание", { hyperlink: false, headingStyleRange: "1-2" }),
  pageBreak(),

  h1("Введение"),
  paragraph("Цель работы — ..."),
  blank(),

  pageBreak(),
  h1("1 Практическая часть"),
  h2("1.1 Постановка задачи"),
  paragraph("Дана матрица ..."),

  formulaMath(["ν = ", mathFraction("48", "13"), " ≈ 3.69"]),

  makeTable([
    ["Критерий", "Значение"],
    ["MM",  "E1"],
    ["BL",  "E2"],
  ], { widths: [4000, 5355] }),

  blank(),
  placeholder("[Рисунок 1 — результат работы программы]"),
  caption("Рисунок 1 – результат работы программы"),

  ...codeBlock(["node compute.js", "// → ν = 48/13"]),
];

const doc = makeDocument([
  makeTitlePage({
    workType: "ИНДИВИДУАЛЬНОЕ ЗАДАНИЕ",
    subtitle: "по теории принятия решений",
    discipline: "по дисциплине: «ТПР»",
    author: "Валиев М. М.",
    group: "КТбо3-7",
    teacher: "Родзин С. И.",
    year: 2026,
  }),
  makeContentSection(children),
]);

saveDocument(doc, "ТПР_ИЗ_Вариант2.docx").then(() => console.log("Done"));
```

---

## Требования к реализации

1. **Нет side-effects при импорте** — библиотека только экспортирует функции, ничего не запускает
2. **Нет жёстко зашитых данных** — никаких "Валиев М. М." внутри библиотеки, всё через параметры
3. **Backward-compatible с build_report.js** — функции `run`, `paragraph`, `centered`, `h1/h2/h3`, `blank`, `pageBreak`, все math-функции должны работать идентично тому, как они работали в старом скрипте
4. **JSDoc на каждой публичной функции** — минимально: `@param` и `@returns`
5. **Без лишних зависимостей** — только `docx` и встроенный `fs`
6. **`TableOfContents` реэкспортировать** из `docx` — чтобы пользователь не делал отдельный `require('docx')`

---

## Чего НЕ делать

- Не создавать отдельные файлы на каждую секцию — весь код в одном `docx-gost.js`
- Не реализовывать `formula(number)` со switch-case — это специфично для каждого отчёта
- Не включать данные варианта (матрицы, таблицы издательств и т.д.)
- Не устанавливать зависимости автоматически

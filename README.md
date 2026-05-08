# docx-gost

Локальная CommonJS-библиотека для сборки `.docx`-отчетов по ГОСТ на базе пакета `docx`.

## Установка зависимости

```bash
npm install docx
```

## Минимальный пример

```js
const {
  h1,
  h2,
  blank,
  centered,
  paragraph,
  codeLine,
  codeBlock,
  caption,
  placeholder,
  imageBlock,
  formulaMath,
  mathFraction,
  mathSub,
  makeTable,
  makeTitlePage,
  makeContentSection,
  makeDocument,
  saveDocument,
  pageBreak,
  TableOfContents,
} = require("../lib/docx-gost");

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
    ["MM", "E1"],
    ["BL", "E2"],
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

`TableOfContents` реэкспортируется из `docx`, поэтому отдельный `require("docx")` для оглавления не нужен.

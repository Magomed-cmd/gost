import { createDocxGost } from "../docx-gost";
import { lintChildren } from "../lint";

const g = createDocxGost();
const { paragraph, h1, h2, blank, pageBreak, imageBlock, placeholder } = g;

// Минимальный PNG-буфер для imageBlock
const fakePng = (() => {
  const b = Buffer.alloc(24);
  b.writeUInt32BE(800, 16);
  b.writeUInt32BE(600, 20);
  return b;
})();

const figures = g.createCaptionCounter("Рисунок");

// ── Утилита ────────────────────────────────────────────────────────────────────

function codes(children: ReturnType<typeof paragraph>[]): string[] {
  return lintChildren(children).map(i => i.code);
}

// ── Пустые страницы ────────────────────────────────────────────────────────────

describe("empty-page-start", () => {
  test("pageBreak первым элементом → warn", () => {
    expect(codes([pageBreak(), paragraph("Текст")])).toContain("empty-page-start");
  });

  test("paragraph первым → нет warn", () => {
    expect(codes([paragraph("Текст"), pageBreak()])).not.toContain("empty-page-start");
  });
});

describe("empty-page-end", () => {
  test("pageBreak последним элементом → warn", () => {
    expect(codes([paragraph("Текст"), pageBreak()])).toContain("empty-page-end");
  });

  test("paragraph последним → нет warn", () => {
    expect(codes([pageBreak(), paragraph("Текст")])).not.toContain("empty-page-end");
  });
});

describe("consecutive-page-breaks", () => {
  test("два pageBreak подряд → warn", () => {
    expect(codes([paragraph("A"), pageBreak(), pageBreak(), paragraph("B")])).toContain("consecutive-page-breaks");
  });

  test("pageBreak с параграфом между → нет warn", () => {
    expect(codes([pageBreak(), paragraph("A"), pageBreak()])).not.toContain("consecutive-page-breaks");
  });
});

describe("page-break-blank-page-break", () => {
  test("pageBreak → blank → pageBreak → warn", () => {
    expect(codes([paragraph("A"), pageBreak(), blank(), pageBreak(), paragraph("B")])).toContain("page-break-blank-page-break");
  });

  test("pageBreak → несколько blank → pageBreak → warn", () => {
    expect(codes([paragraph("A"), pageBreak(), blank(), blank(), blank(), pageBreak(), paragraph("B")])).toContain("page-break-blank-page-break");
  });

  test("pageBreak → paragraph → pageBreak → нет warn", () => {
    expect(codes([paragraph("A"), pageBreak(), paragraph("B"), pageBreak(), paragraph("C")])).not.toContain("page-break-blank-page-break");
  });
});

// ── Заголовки ──────────────────────────────────────────────────────────────────

describe("heading-before-page-break", () => {
  test("h1 прямо перед pageBreak → warn", () => {
    expect(codes([paragraph("A"), h1("Раздел"), pageBreak(), paragraph("B")])).toContain("heading-before-page-break");
  });

  test("h2 прямо перед pageBreak → warn", () => {
    expect(codes([h2("Подраздел"), pageBreak()])).toContain("heading-before-page-break");
  });

  test("h1 с параграфом перед pageBreak → нет warn", () => {
    expect(codes([h1("Раздел"), paragraph("Текст"), pageBreak()])).not.toContain("heading-before-page-break");
  });
});

// ── Лишние элементы ────────────────────────────────────────────────────────────

describe("blank-before-page-break", () => {
  test("blank перед pageBreak → info", () => {
    expect(codes([paragraph("A"), blank(), pageBreak(), paragraph("B")])).toContain("blank-before-page-break");
  });

  test("paragraph перед pageBreak → нет info", () => {
    expect(codes([paragraph("A"), pageBreak(), paragraph("B")])).not.toContain("blank-before-page-break");
  });
});

describe("too-many-blanks", () => {
  test("3 blank подряд → info", () => {
    expect(codes([blank(), blank(), blank()])).toContain("too-many-blanks");
  });

  test("2 blank подряд → нет info", () => {
    expect(codes([blank(), blank()])).not.toContain("too-many-blanks");
  });

  test("4 blank подряд → одно предупреждение (не 4)", () => {
    const issues = lintChildren([blank(), blank(), blank(), blank()]);
    expect(issues.filter(i => i.code === "too-many-blanks")).toHaveLength(1);
  });
});

// ── ГОСТ-нарушения ─────────────────────────────────────────────────────────────

describe("image-without-caption", () => {
  test("image без caption следом → warn", () => {
    expect(codes([imageBlock(fakePng, 400, 300), paragraph("Текст")])).toContain("image-without-caption");
  });

  test("image с caption следом → нет warn", () => {
    const cap = figures.caption("описание");
    expect(codes([imageBlock(fakePng, 400, 300), cap])).not.toContain("image-without-caption");
  });
});

describe("placeholder-left", () => {
  test("placeholder в children → warn", () => {
    expect(codes([paragraph("A"), placeholder("[Рисунок]"), paragraph("B")])).toContain("placeholder-left");
  });

  test("нет placeholder → нет warn", () => {
    expect(codes([paragraph("A"), paragraph("B")])).not.toContain("placeholder-left");
  });
});

// ── Разрыв рисунок/подпись ────────────────────────────────────────────────────

describe("image-caption-split", () => {
  test("image → pageBreak → caption → warn", () => {
    const cap = figures.caption("описание");
    expect(codes([imageBlock(fakePng, 400, 300), pageBreak(), cap])).toContain("image-caption-split");
  });

  test("image → blank → pageBreak → caption → warn", () => {
    const cap = figures.caption("описание");
    expect(codes([imageBlock(fakePng, 400, 300), blank(), pageBreak(), cap])).toContain("image-caption-split");
  });

  test("image → caption (без pageBreak) → нет warn", () => {
    const cap = figures.caption("описание");
    expect(codes([imageBlock(fakePng, 400, 300), cap])).not.toContain("image-caption-split");
  });
});

describe("table-caption-split", () => {
  test("tableCaption → pageBreak → нет таблицы → warn", () => {
    const tables = g.createCaptionCounter("Таблица", { table: true });
    const cap = tables.caption("данные");
    expect(codes([cap, pageBreak(), paragraph("текст")])).toContain("table-caption-split");
  });

  test("tableCaption → blank → pageBreak → warn", () => {
    const tables = g.createCaptionCounter("Таблица", { table: true });
    const cap = tables.caption("данные");
    expect(codes([cap, blank(), pageBreak()])).toContain("table-caption-split");
  });

  test("tableCaption → table (без pageBreak) → нет warn", () => {
    const tables = g.createCaptionCounter("Таблица", { table: true });
    const cap = tables.caption("данные");
    const tbl = g.makeTable([["A", "B"]]);
    expect(codes([cap, tbl as unknown as ReturnType<typeof paragraph>])).not.toContain("table-caption-split");
  });
});

// ── Чистый документ ───────────────────────────────────────────────────────────

describe("чистый документ", () => {
  test("нормальная структура не даёт предупреждений", () => {
    const children = [
      h1("Введение"),
      paragraph("Текст введения."),
      pageBreak(),
      h1("Раздел 1"),
      paragraph("Текст раздела."),
    ];
    expect(lintChildren(children)).toHaveLength(0);
  });
});

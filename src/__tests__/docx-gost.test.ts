import { createDocxGost } from "../docx-gost";
import { Paragraph } from "docx";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";

describe("createDocxGost", () => {
  const g = createDocxGost();

  test("возвращает инстанс со всеми методами", () => {
    const methods = [
      "paragraph", "centered", "h1", "h2", "h3",
      "blank", "pageBreak", "caption", "tableCaption",
      "createCaptionCounter", "placeholder", "imageBlock",
      "codeLine", "codeBlock", "formulaMath", "formulaInline",
      "makeTable", "makeTitlePage", "makeContentSection",
      "makeDocument", "saveDocument", "diagramBlock", "codeFile",
    ];
    methods.forEach((m) => expect(typeof (g as unknown as Record<string, unknown>)[m]).toBe("function"));
  });

  test("paragraph() возвращает Paragraph", () => {
    expect(g.paragraph("Текст")).toBeInstanceOf(Paragraph);
  });

  test("h1() возвращает Paragraph", () => {
    expect(g.h1("Заголовок")).toBeInstanceOf(Paragraph);
  });

  test("blank() возвращает Paragraph", () => {
    expect(g.blank()).toBeInstanceOf(Paragraph);
  });

  test("pageBreak() возвращает Paragraph", () => {
    expect(g.pageBreak()).toBeInstanceOf(Paragraph);
  });

  test("codeBlock() возвращает массив Paragraph", () => {
    const result = g.codeBlock(["line1", "line2", "line3"]);
    expect(Array.isArray(result)).toBe(true);
    expect(result).toHaveLength(3);
    result.forEach((p) => expect(p).toBeInstanceOf(Paragraph));
  });

  describe("createCaptionCounter", () => {
    test("peek() возвращает текущий номер без инкремента", () => {
      const c = g.createCaptionCounter("Рисунок");
      expect(c.peek()).toBe(1);
      expect(c.peek()).toBe(1); // не меняется
    });

    test("caption() инкрементирует счётчик", () => {
      const c = g.createCaptionCounter("Рисунок");
      expect(c.peek()).toBe(1);
      c.caption("первый рисунок");
      expect(c.peek()).toBe(2);
      c.caption("второй рисунок");
      expect(c.peek()).toBe(3);
    });

    test("caption() возвращает Paragraph", () => {
      const c = g.createCaptionCounter("Таблица");
      expect(c.caption("данные")).toBeInstanceOf(Paragraph);
    });

    test("start: 5 начинает с 5", () => {
      const c = g.createCaptionCounter("Рисунок", { start: 5 });
      expect(c.peek()).toBe(5);
    });
  });

  // Минимальный валидный PNG-буфер: сигнатура + IHDR с реальными размерами
  const fakePng = (() => {
    const b = Buffer.alloc(24);
    // PNG magic bytes: \x89PNG\r\n\x1a\n
    b[0] = 0x89; b[1] = 0x50; b[2] = 0x4e; b[3] = 0x47;
    b[4] = 0x0d; b[5] = 0x0a; b[6] = 0x1a; b[7] = 0x0a;
    b.writeUInt32BE(800, 16);
    b.writeUInt32BE(600, 20);
    return b;
  })();

  describe("imageBlock", () => {
    test("imageBlock(src, width, height) — явные размеры возвращает Paragraph", () => {
      expect(g.imageBlock(fakePng, 600, 400)).toBeInstanceOf(Paragraph);
    });

    test("imageBlock(src) — авто-размер из PNG-байт возвращает Paragraph", () => {
      expect(g.imageBlock(fakePng)).toBeInstanceOf(Paragraph);
    });

    test("imageBlock(src, { maxWidth: 300 }) — авто-размер с кастомной шириной", () => {
      // 800×600 с maxWidth=300 → 300×225
      expect(g.imageBlock(fakePng, { maxWidth: 300 })).toBeInstanceOf(Paragraph);
    });

    test("imageBlock(src) без PNG-сигнатуры — бросает ошибку", () => {
      const notPng = Buffer.alloc(24);
      expect(() => g.imageBlock(notPng)).toThrow(/PNG/);
    });
  });

  describe("codeFile", () => {
    let tmpFile: string;

    beforeEach(() => {
      tmpFile = path.join(os.tmpdir(), `gost-test-${Date.now()}.txt`);
    });

    afterEach(() => {
      if (fs.existsSync(tmpFile)) fs.unlinkSync(tmpFile);
    });

    test("возвращает массив Paragraph по числу строк файла", () => {
      fs.writeFileSync(tmpFile, "line1\nline2\nline3");
      const result = g.codeFile(tmpFile);
      expect(result).toHaveLength(3);
      result.forEach((p) => expect(p).toBeInstanceOf(Paragraph));
    });

    test("файл, оканчивающийся на \\n, не даёт лишнюю пустую строку", () => {
      fs.writeFileSync(tmpFile, "line1\nline2\n");
      expect(g.codeFile(tmpFile)).toHaveLength(2);
    });

    test("Windows CRLF нормализуется в LF", () => {
      fs.writeFileSync(tmpFile, "line1\r\nline2\r\nline3\r\n");
      expect(g.codeFile(tmpFile)).toHaveLength(3);
    });

    test("табы заменяются на 2 пробела", () => {
      fs.writeFileSync(tmpFile, "\tindented");
      // Просто проверяем что не бросает исключение и возвращает Paragraph
      expect(g.codeFile(tmpFile)[0]).toBeInstanceOf(Paragraph);
    });

    test("пустой файл возвращает пустой массив", () => {
      fs.writeFileSync(tmpFile, "");
      expect(g.codeFile(tmpFile)).toHaveLength(0);
    });
  });

  describe("makeTable", () => {
    test("не бросает исключение для простой таблицы", () => {
      expect(() =>
        g.makeTable([
          ["Заголовок 1", "Заголовок 2"],
          ["Значение 1", "Значение 2"],
        ])
      ).not.toThrow();
    });
  });

  describe("makeDocument", () => {
    test("не бросает исключение при сборке документа", () => {
      const titlePage = g.makeTitlePage({
        workType: "ОТЧЁТ",
        subtitle: ["о выполнении работы №1"],
        discipline: "по дисциплине: «Тест»",
        group: "КТбо3-7",
        author: "Валиев М. М.",
        teacher: "Лутай В. Н.",
        year: 2026,
      });
      const content = g.makeContentSection([g.h1("Раздел"), g.paragraph("Текст.")]);
      expect(() => g.makeDocument([titlePage, content])).not.toThrow();
    });
  });
});

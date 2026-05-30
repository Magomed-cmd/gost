"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const docx_gost_1 = require("../docx-gost");
const docx_1 = require("docx");
describe("createDocxGost", () => {
    const g = (0, docx_gost_1.createDocxGost)();
    test("возвращает инстанс со всеми методами", () => {
        const methods = [
            "paragraph", "centered", "h1", "h2", "h3",
            "blank", "pageBreak", "caption", "tableCaption",
            "createCaptionCounter", "placeholder", "imageBlock",
            "codeLine", "codeBlock", "formulaMath", "formulaInline",
            "makeTable", "makeTitlePage", "makeContentSection",
            "makeDocument", "saveDocument", "diagramBlock",
        ];
        methods.forEach((m) => expect(typeof g[m]).toBe("function"));
    });
    test("paragraph() возвращает Paragraph", () => {
        expect(g.paragraph("Текст")).toBeInstanceOf(docx_1.Paragraph);
    });
    test("h1() возвращает Paragraph", () => {
        expect(g.h1("Заголовок")).toBeInstanceOf(docx_1.Paragraph);
    });
    test("blank() возвращает Paragraph", () => {
        expect(g.blank()).toBeInstanceOf(docx_1.Paragraph);
    });
    test("pageBreak() возвращает Paragraph", () => {
        expect(g.pageBreak()).toBeInstanceOf(docx_1.Paragraph);
    });
    test("codeBlock() возвращает массив Paragraph", () => {
        const result = g.codeBlock(["line1", "line2", "line3"]);
        expect(Array.isArray(result)).toBe(true);
        expect(result).toHaveLength(3);
        result.forEach((p) => expect(p).toBeInstanceOf(docx_1.Paragraph));
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
            expect(c.caption("данные")).toBeInstanceOf(docx_1.Paragraph);
        });
        test("start: 5 начинает с 5", () => {
            const c = g.createCaptionCounter("Рисунок", { start: 5 });
            expect(c.peek()).toBe(5);
        });
    });
    describe("makeTable", () => {
        test("не бросает исключение для простой таблицы", () => {
            expect(() => g.makeTable([
                ["Заголовок 1", "Заголовок 2"],
                ["Значение 1", "Значение 2"],
            ])).not.toThrow();
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
//# sourceMappingURL=docx-gost.test.js.map
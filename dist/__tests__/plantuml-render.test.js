"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const plantuml_render_1 = require("../plantuml-render");
describe("wrapWithSkin", () => {
    test("добавляет базовые skinparam после @startuml", () => {
        const src = "@startuml\nA --> B\n@enduml";
        const result = (0, plantuml_render_1.wrapWithSkin)(src);
        expect(result).toContain("skinparam defaultFontName");
        expect(result).toContain("Times New Roman");
        expect(result).toContain("skinparam dpi 150");
        expect(result).toContain("A --> B");
        expect(result.indexOf("skinparam")).toBeLessThan(result.indexOf("A --> B"));
    });
    test("использует переданный dpi", () => {
        const src = "@startuml\nA --> B\n@enduml";
        const result = (0, plantuml_render_1.wrapWithSkin)(src, 300);
        expect(result).toContain("skinparam dpi 300");
    });
    test("opts.skinparams переопределяет базовые значения", () => {
        const src = "@startuml\nA --> B\n@enduml";
        const result = (0, plantuml_render_1.wrapWithSkin)(src, 150, { skinparams: { defaultFontName: "Arial" } });
        expect(result).toContain("skinparam defaultFontName Arial");
        expect(result).not.toContain("Times New Roman");
    });
    test("opts.skinparams добавляет новые ключи", () => {
        const src = "@startuml\nA --> B\n@enduml";
        const result = (0, plantuml_render_1.wrapWithSkin)(src, 150, { skinparams: { backgroundColor: "white" } });
        expect(result).toContain("skinparam backgroundColor white");
    });
    test("не ломает структуру @startuml / @enduml", () => {
        const src = "@startuml\nA --> B\n@enduml";
        const result = (0, plantuml_render_1.wrapWithSkin)(src);
        expect(result.startsWith("@startuml")).toBe(true);
        expect(result.endsWith("@enduml")).toBe(true);
    });
});
//# sourceMappingURL=plantuml-render.test.js.map
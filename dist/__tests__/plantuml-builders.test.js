"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const plantuml_gost_1 = require("../plantuml-gost");
// ── Базовый строитель ──────────────────────────────────────────────────────────
describe("DiagramBuilder", () => {
    test("build() оборачивает в @startuml / @enduml", () => {
        const puml = new plantuml_gost_1.DiagramBuilder().add("A --> B").build();
        expect(puml).toMatch(/^@startuml/);
        expect(puml).toMatch(/@enduml$/);
        expect(puml).toContain("A --> B");
    });
    test("add() поддерживает chaining", () => {
        const puml = new plantuml_gost_1.DiagramBuilder().add("A --> B").add("B --> C").build();
        expect(puml).toContain("A --> B");
        expect(puml).toContain("B --> C");
    });
    test("blank() добавляет пустую строку", () => {
        const puml = new plantuml_gost_1.DiagramBuilder().add("A").blank().add("B").build();
        expect(puml).toContain("A\n\nB");
    });
});
// ── UseCaseBuilder ─────────────────────────────────────────────────────────────
describe("UseCaseBuilder", () => {
    test("actor() без alias", () => {
        const puml = new plantuml_gost_1.UseCaseBuilder().actor("Водитель").build();
        expect(puml).toContain('actor "Водитель"');
    });
    test("actor() с alias", () => {
        const puml = new plantuml_gost_1.UseCaseBuilder().actor("Водитель", "D").build();
        expect(puml).toContain('actor "Водитель" as D');
    });
    test("usecase() без alias", () => {
        const puml = new plantuml_gost_1.UseCaseBuilder().usecase("Сообщить о поломке").build();
        expect(puml).toContain('usecase "Сообщить о поломке"');
    });
    test("arrow() без label", () => {
        const puml = new plantuml_gost_1.UseCaseBuilder().arrow("A", "B").build();
        expect(puml).toContain("A --> B");
    });
    test("arrow() с label", () => {
        const puml = new plantuml_gost_1.UseCaseBuilder().arrow("A", "B", "запрос").build();
        expect(puml).toContain("A --> B : запрос");
    });
    test("extend() добавляет <<extend>>", () => {
        const puml = new plantuml_gost_1.UseCaseBuilder().extend("UC1", "UC2").build();
        expect(puml).toContain("<<extend>>");
    });
    test("include() добавляет <<include>>", () => {
        const puml = new plantuml_gost_1.UseCaseBuilder().include("UC1", "UC2").build();
        expect(puml).toContain("<<include>>");
    });
    test("package() оборачивает в rectangle", () => {
        const puml = new plantuml_gost_1.UseCaseBuilder()
            .package("Система", (b) => b.usecase("Действие"))
            .build();
        expect(puml).toContain('rectangle "Система" {');
        expect(puml).toContain('usecase "Действие"');
        expect(puml).toContain("}");
    });
});
// ── ClassDiagramBuilder ────────────────────────────────────────────────────────
describe("ClassDiagramBuilder", () => {
    test("build() добавляет hide circle", () => {
        const puml = new plantuml_gost_1.ClassDiagramBuilder().build();
        expect(puml).toContain("hide circle");
    });
    test("class() без членов", () => {
        const puml = new plantuml_gost_1.ClassDiagramBuilder().class("Tram").build();
        expect(puml).toContain("class Tram");
    });
    test("class() с членами", () => {
        const puml = new plantuml_gost_1.ClassDiagramBuilder()
            .class("Tram", ["+ getId(): String", "- state: TramState"])
            .build();
        expect(puml).toContain("class Tram {");
        expect(puml).toContain("+ getId(): String");
        expect(puml).toContain("- state: TramState");
        expect(puml).toContain("}");
    });
    test("interface() генерирует interface", () => {
        const puml = new plantuml_gost_1.ClassDiagramBuilder().interface("RepairStrategy").build();
        expect(puml).toContain("interface RepairStrategy");
    });
    test("enum() с значениями", () => {
        const puml = new plantuml_gost_1.ClassDiagramBuilder().enum("TramState", ["ON_LINE", "BROKEN"]).build();
        expect(puml).toContain("enum TramState {");
        expect(puml).toContain("ON_LINE");
        expect(puml).toContain("BROKEN");
    });
    test("extends() генерирует --|>", () => {
        const puml = new plantuml_gost_1.ClassDiagramBuilder().extends("Child", "Parent").build();
        expect(puml).toContain("Child --|> Parent");
    });
    test("implements() генерирует ..|>", () => {
        const puml = new plantuml_gost_1.ClassDiagramBuilder().implements("Cls", "Iface").build();
        expect(puml).toContain("Cls ..|> Iface");
    });
    test("dependency() генерирует ..>", () => {
        const puml = new plantuml_gost_1.ClassDiagramBuilder().dependency("A", "B", "использует").build();
        expect(puml).toContain("A ..> B : использует");
    });
    test("composition() генерирует *--", () => {
        const puml = new plantuml_gost_1.ClassDiagramBuilder().composition("A", "B").build();
        expect(puml).toContain("A *-- B");
    });
});
// ── SequenceBuilder ────────────────────────────────────────────────────────────
describe("SequenceBuilder", () => {
    test("participant() без alias", () => {
        const puml = new plantuml_gost_1.SequenceBuilder().participant("Диспетчер").build();
        expect(puml).toContain('participant "Диспетчер"');
    });
    test("arrow() генерирует ->", () => {
        const puml = new plantuml_gost_1.SequenceBuilder().arrow("A", "B", "запрос").build();
        expect(puml).toContain("A -> B : запрос");
    });
    test("return() генерирует -->", () => {
        const puml = new plantuml_gost_1.SequenceBuilder().return("B", "A", "ответ").build();
        expect(puml).toContain("B --> A : ответ");
    });
    test("activate/deactivate", () => {
        const puml = new plantuml_gost_1.SequenceBuilder().activate("A").deactivate("A").build();
        expect(puml).toContain("activate A");
        expect(puml).toContain("deactivate A");
    });
    test("divider() с текстом", () => {
        const puml = new plantuml_gost_1.SequenceBuilder().divider("Ремонт").build();
        expect(puml).toContain("== Ремонт ==");
    });
    test("group() оборачивает в group/end", () => {
        const puml = new plantuml_gost_1.SequenceBuilder()
            .group("alt успех", (b) => b.arrow("A", "B", "ok"))
            .build();
        expect(puml).toContain("group alt успех");
        expect(puml).toContain("end");
    });
});
// ── StateBuilder ───────────────────────────────────────────────────────────────
describe("StateBuilder", () => {
    test("initial() генерирует [*] -->", () => {
        const puml = new plantuml_gost_1.StateBuilder().initial("Active").build();
        expect(puml).toContain("[*] --> Active");
    });
    test("final() генерирует --> [*]", () => {
        const puml = new plantuml_gost_1.StateBuilder().final("Done").build();
        expect(puml).toContain("Done --> [*]");
    });
    test("transition() без label и direction", () => {
        const puml = new plantuml_gost_1.StateBuilder().transition("A", "B").build();
        expect(puml).toContain("A --> B");
    });
    test("transition() с направлением -down->", () => {
        const puml = new plantuml_gost_1.StateBuilder().transition("A", "B", "событие", "down").build();
        expect(puml).toContain("A -down-> B : событие");
    });
    test("state() с label", () => {
        const puml = new plantuml_gost_1.StateBuilder().state("S1", "На линии").build();
        expect(puml).toContain('state "На линии" as S1');
    });
});
//# sourceMappingURL=plantuml-builders.test.js.map
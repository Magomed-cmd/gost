import {
  DiagramBuilder,
  UseCaseBuilder,
  ClassDiagramBuilder,
  SequenceBuilder,
  StateBuilder,
  ActivityBuilder,
  ComponentBuilder,
  DeploymentBuilder,
  ObjectBuilder,
  ERBuilder,
  ER,
  TimingBuilder,
  MindMapBuilder,
  GanttBuilder,
} from "../plantuml-gost";

// ── Базовый строитель ──────────────────────────────────────────────────────────

describe("DiagramBuilder", () => {
  test("build() оборачивает в @startuml / @enduml", () => {
    const puml = new DiagramBuilder().add("A --> B").build();
    expect(puml).toMatch(/^@startuml/);
    expect(puml).toMatch(/@enduml$/);
    expect(puml).toContain("A --> B");
  });

  test("add() поддерживает chaining", () => {
    const puml = new DiagramBuilder().add("A --> B").add("B --> C").build();
    expect(puml).toContain("A --> B");
    expect(puml).toContain("B --> C");
  });

  test("blank() добавляет пустую строку", () => {
    const puml = new DiagramBuilder().add("A").blank().add("B").build();
    expect(puml).toContain("A\n\nB");
  });
});

// ── UseCaseBuilder ─────────────────────────────────────────────────────────────

describe("UseCaseBuilder", () => {
  test("actor() без alias", () => {
    const puml = new UseCaseBuilder().actor("Водитель").build();
    expect(puml).toContain('actor "Водитель"');
  });

  test("actor() с alias", () => {
    const puml = new UseCaseBuilder().actor("Водитель", "D").build();
    expect(puml).toContain('actor "Водитель" as D');
  });

  test("usecase() без alias", () => {
    const puml = new UseCaseBuilder().usecase("Сообщить о поломке").build();
    expect(puml).toContain('usecase "Сообщить о поломке"');
  });

  test("arrow() без label", () => {
    const puml = new UseCaseBuilder().arrow("A", "B").build();
    expect(puml).toContain("A --> B");
  });

  test("arrow() с label", () => {
    const puml = new UseCaseBuilder().arrow("A", "B", "запрос").build();
    expect(puml).toContain("A --> B : запрос");
  });

  test("extend() добавляет <<extend>>", () => {
    const puml = new UseCaseBuilder().extend("UC1", "UC2").build();
    expect(puml).toContain("<<extend>>");
  });

  test("include() добавляет <<include>>", () => {
    const puml = new UseCaseBuilder().include("UC1", "UC2").build();
    expect(puml).toContain("<<include>>");
  });

  test("package() оборачивает в rectangle", () => {
    const puml = new UseCaseBuilder()
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
    const puml = new ClassDiagramBuilder().build();
    expect(puml).toContain("hide circle");
  });

  test("class() без членов", () => {
    const puml = new ClassDiagramBuilder().class("Tram").build();
    expect(puml).toContain("class Tram");
  });

  test("class() с членами", () => {
    const puml = new ClassDiagramBuilder()
      .class("Tram", ["+ getId(): String", "- state: TramState"])
      .build();
    expect(puml).toContain("class Tram {");
    expect(puml).toContain("+ getId(): String");
    expect(puml).toContain("- state: TramState");
    expect(puml).toContain("}");
  });

  test("interface() генерирует interface", () => {
    const puml = new ClassDiagramBuilder().interface("RepairStrategy").build();
    expect(puml).toContain("interface RepairStrategy");
  });

  test("enum() с значениями", () => {
    const puml = new ClassDiagramBuilder().enum("TramState", ["ON_LINE", "BROKEN"]).build();
    expect(puml).toContain("enum TramState {");
    expect(puml).toContain("ON_LINE");
    expect(puml).toContain("BROKEN");
  });

  test("extends() генерирует --|>", () => {
    const puml = new ClassDiagramBuilder().extends("Child", "Parent").build();
    expect(puml).toContain("Child --|> Parent");
  });

  test("implements() генерирует ..|>", () => {
    const puml = new ClassDiagramBuilder().implements("Cls", "Iface").build();
    expect(puml).toContain("Cls ..|> Iface");
  });

  test("dependency() генерирует ..>", () => {
    const puml = new ClassDiagramBuilder().dependency("A", "B", "использует").build();
    expect(puml).toContain("A ..> B : использует");
  });

  test("composition() генерирует *--", () => {
    const puml = new ClassDiagramBuilder().composition("A", "B").build();
    expect(puml).toContain("A *-- B");
  });
});

// ── SequenceBuilder ────────────────────────────────────────────────────────────

describe("SequenceBuilder", () => {
  test("participant() без alias", () => {
    const puml = new SequenceBuilder().participant("Диспетчер").build();
    expect(puml).toContain('participant "Диспетчер"');
  });

  test("arrow() генерирует ->", () => {
    const puml = new SequenceBuilder().arrow("A", "B", "запрос").build();
    expect(puml).toContain("A -> B : запрос");
  });

  test("return() генерирует -->", () => {
    const puml = new SequenceBuilder().return("B", "A", "ответ").build();
    expect(puml).toContain("B --> A : ответ");
  });

  test("activate/deactivate", () => {
    const puml = new SequenceBuilder().activate("A").deactivate("A").build();
    expect(puml).toContain("activate A");
    expect(puml).toContain("deactivate A");
  });

  test("divider() с текстом", () => {
    const puml = new SequenceBuilder().divider("Ремонт").build();
    expect(puml).toContain("== Ремонт ==");
  });

  test("group() оборачивает в group/end", () => {
    const puml = new SequenceBuilder()
      .group("alt успех", (b) => b.arrow("A", "B", "ok"))
      .build();
    expect(puml).toContain("group alt успех");
    expect(puml).toContain("end");
  });
});

// ── StateBuilder ───────────────────────────────────────────────────────────────

describe("StateBuilder", () => {
  test("initial() генерирует [*] -->", () => {
    const puml = new StateBuilder().initial("Active").build();
    expect(puml).toContain("[*] --> Active");
  });

  test("final() генерирует --> [*]", () => {
    const puml = new StateBuilder().final("Done").build();
    expect(puml).toContain("Done --> [*]");
  });

  test("transition() без label и direction", () => {
    const puml = new StateBuilder().transition("A", "B").build();
    expect(puml).toContain("A --> B");
  });

  test("transition() с направлением -down->", () => {
    const puml = new StateBuilder().transition("A", "B", "событие", "down").build();
    expect(puml).toContain("A -down-> B : событие");
  });

  test("state() с label", () => {
    const puml = new StateBuilder().state("S1", "На линии").build();
    expect(puml).toContain('state "На линии" as S1');
  });
});

// ── ActivityBuilder ────────────────────────────────────────────────────────────

describe("ActivityBuilder", () => {
  test("start/stop генерирует ключевые слова", () => {
    const puml = new ActivityBuilder().start().stop().build();
    expect(puml).toContain("start");
    expect(puml).toContain("stop");
  });

  test("action() оборачивает в :...;", () => {
    const puml = new ActivityBuilder().action("Обработать запрос").build();
    expect(puml).toContain(":Обработать запрос;");
  });

  test("if() без else — генерирует if/endif", () => {
    const puml = new ActivityBuilder()
      .if("условие", (b) => b.action("да"))
      .build();
    expect(puml).toContain("if (условие?) then (yes)");
    expect(puml).toContain(":да;");
    expect(puml).toContain("endif");
    expect(puml).not.toContain("else");
  });

  test("if() с else — генерирует if/else/endif", () => {
    const puml = new ActivityBuilder()
      .if("условие", (b) => b.action("да"), (b) => b.action("нет"))
      .build();
    expect(puml).toContain("else (no)");
    expect(puml).toContain(":нет;");
  });

  test("while() генерирует while/endwhile", () => {
    const puml = new ActivityBuilder()
      .while("есть задачи", (b) => b.action("обработать"))
      .build();
    expect(puml).toContain("while (есть задачи?)");
    expect(puml).toContain("endwhile");
  });

  test("while() с exitLabel", () => {
    const puml = new ActivityBuilder()
      .while("есть задачи", (b) => b.action("обработать"), "стоп")
      .build();
    expect(puml).toContain("endwhile (стоп)");
  });

  test("fork() генерирует fork/fork again/end fork", () => {
    const puml = new ActivityBuilder()
      .fork(
        (b) => b.action("ветвь 1"),
        (b) => b.action("ветвь 2"),
      )
      .build();
    expect(puml).toContain("fork");
    expect(puml).toContain("fork again");
    expect(puml).toContain("end fork");
    expect(puml).toContain(":ветвь 1;");
    expect(puml).toContain(":ветвь 2;");
  });

  test("partition() оборачивает в swimlane", () => {
    const puml = new ActivityBuilder()
      .partition("Сервер", (b) => b.action("обработать"))
      .build();
    expect(puml).toContain('partition "Сервер" {');
    expect(puml).toContain("}");
  });

  test("note() добавляет заметку", () => {
    const puml = new ActivityBuilder().note("важно", "left").build();
    expect(puml).toContain("note left: важно");
  });
});

// ── ComponentBuilder ───────────────────────────────────────────────────────────

describe("ComponentBuilder", () => {
  test("component() без alias", () => {
    const puml = new ComponentBuilder().component("Frontend").build();
    expect(puml).toContain("component [Frontend]");
  });

  test("component() с alias", () => {
    const puml = new ComponentBuilder().component("Frontend", "FE").build();
    expect(puml).toContain("component [Frontend] as FE");
  });

  test("interface() генерирует interface", () => {
    const puml = new ComponentBuilder().interface("IService").build();
    expect(puml).toContain('interface "IService"');
  });

  test("arrow() с label", () => {
    const puml = new ComponentBuilder().arrow("FE", "BE", "HTTP").build();
    expect(puml).toContain("FE --> BE : HTTP");
  });

  test("link() без label", () => {
    const puml = new ComponentBuilder().link("A", "B").build();
    expect(puml).toContain("A -- B");
  });

  test("package() оборачивает в package {...}", () => {
    const puml = new ComponentBuilder()
      .package("Система", (b) => b.component("Core"))
      .build();
    expect(puml).toContain('package "Система" {');
    expect(puml).toContain("component [Core]");
  });
});

// ── DeploymentBuilder ──────────────────────────────────────────────────────────

describe("DeploymentBuilder", () => {
  test("node() без вложения", () => {
    const puml = new DeploymentBuilder().node("WebServer", "WS").build();
    expect(puml).toContain('node "WebServer" as WS');
  });

  test("node() с вложением генерирует блок", () => {
    const puml = new DeploymentBuilder()
      .node("Server", undefined, (b) => b.component("App"))
      .build();
    expect(puml).toContain('node "Server" {');
    expect(puml).toContain("component [App]");
    expect(puml).toContain("}");
  });

  test("artifact() генерирует artifact", () => {
    const puml = new DeploymentBuilder().artifact("app.jar").build();
    expect(puml).toContain('artifact "app.jar"');
  });

  test("cloud() с вложением", () => {
    const puml = new DeploymentBuilder()
      .cloud("AWS", (b) => b.node("EC2"))
      .build();
    expect(puml).toContain('cloud "AWS" {');
    expect(puml).toContain('node "EC2"');
  });

  test("arrow() и link()", () => {
    const puml = new DeploymentBuilder().arrow("A", "B").link("B", "C").build();
    expect(puml).toContain("A --> B");
    expect(puml).toContain("B -- C");
  });
});

// ── ObjectBuilder ──────────────────────────────────────────────────────────────

describe("ObjectBuilder", () => {
  test("object() без полей — простое имя", () => {
    const puml = new ObjectBuilder().object("tram1").build();
    expect(puml).toContain("object tram1");
  });

  test("object() с полями", () => {
    const puml = new ObjectBuilder()
      .object("tram1", undefined, { id: 1, state: '"working"' })
      .build();
    expect(puml).toContain("object tram1 {");
    expect(puml).toContain("id = 1");
    expect(puml).toContain('state = "working"');
  });

  test("object() с именем содержащим двоеточие — автоматически кавычки + alias", () => {
    const puml = new ObjectBuilder()
      .object("tram1 : Tram", "tram1")
      .build();
    expect(puml).toContain('object "tram1 : Tram" as tram1');
  });

  test("link() со стрелкой", () => {
    const puml = new ObjectBuilder().link("A", "B", "has").build();
    expect(puml).toContain("A --> B : has");
  });
});

// ── ERBuilder ──────────────────────────────────────────────────────────────────

describe("ERBuilder", () => {
  test("entity() без обычных полей", () => {
    const puml = new ERBuilder().entity("User", ["id : int <<PK>>"]).build();
    expect(puml).toContain('entity "User" {');
    expect(puml).toContain("* id : int <<PK>>");
    expect(puml).not.toContain("--");
  });

  test("entity() с PK и обычными полями добавляет разделитель --", () => {
    const puml = new ERBuilder()
      .entity("User", ["id : int"], ["name : varchar"])
      .build();
    expect(puml).toContain("* id : int");
    expect(puml).toContain("  --");
    expect(puml).toContain("  name : varchar");
  });

  test("relation() с константами ER.*", () => {
    const puml = new ERBuilder()
      .relation("User", ER.ONE_TO_MANY, "Order", "places")
      .build();
    expect(puml).toContain("User ||--o{ Order : places");
  });

  test("relation() без label", () => {
    const puml = new ERBuilder()
      .relation("A", ER.MANY_TO_MANY, "B")
      .build();
    expect(puml).toContain("A }o--o{ B");
  });

  test("ER константы — корректный PlantUML синтаксис", () => {
    expect(ER.ONE_TO_ONE).toBe("||--||");
    expect(ER.ONE_TO_MANY).toBe("||--o{");
    expect(ER.ONE_TO_MANY_REQ).toBe("||--|{");
    expect(ER.OPT_TO_MANY).toBe("|o--o{");
    expect(ER.MANY_TO_MANY).toBe("}o--o{");
  });
});

// ── TimingBuilder ──────────────────────────────────────────────────────────────

describe("TimingBuilder", () => {
  test("concise() и robust() объявляют участников", () => {
    const puml = new TimingBuilder()
      .concise("User", "U")
      .robust("System", "S")
      .build();
    expect(puml).toContain('concise "User" as U');
    expect(puml).toContain('robust "System" as S');
  });

  test("at() устанавливает временну́ю точку", () => {
    const puml = new TimingBuilder().at(100).build();
    expect(puml).toContain("@100");
  });

  test("state() устанавливает состояние", () => {
    const puml = new TimingBuilder().state("U", "Idle").build();
    expect(puml).toContain("U is Idle");
  });

  test("arrow() генерирует сообщение", () => {
    const puml = new TimingBuilder().arrow("U", "S", "запрос").build();
    expect(puml).toContain("U -> S : запрос");
  });

  test("highlight() с label и цветом", () => {
    const puml = new TimingBuilder().highlight(0, 100, "#Gold", "фаза 1").build();
    expect(puml).toContain("highlight 0 to 100 #Gold : фаза 1");
  });

  test("highlight() без label", () => {
    const puml = new TimingBuilder().highlight(0, 100).build();
    expect(puml).toContain("highlight 0 to 100 #LightYellow");
  });

  test("clock() добавляет тактовый сигнал", () => {
    const puml = new TimingBuilder().clock("CLK", 50).build();
    expect(puml).toContain('clock "CLK" with period 50');
  });
});

// ── MindMapBuilder ─────────────────────────────────────────────────────────────

describe("MindMapBuilder", () => {
  test("build() использует @startmindmap / @endmindmap", () => {
    const puml = new MindMapBuilder().node(1, "Корень").build();
    expect(puml).toMatch(/^@startmindmap/);
    expect(puml).toMatch(/@endmindmap$/);
  });

  test("node(1) — корень со звёздочкой", () => {
    const puml = new MindMapBuilder().node(1, "Root").build();
    expect(puml).toContain("* Root");
  });

  test("node(2) — ветвь с двумя звёздочками", () => {
    const puml = new MindMapBuilder().node(2, "Ветвь").build();
    expect(puml).toContain("** Ветвь");
  });

  test("node(2, text, 'left') — левая ветвь с _", () => {
    const puml = new MindMapBuilder().node(2, "Левая", "left").build();
    expect(puml).toContain("**_ Левая");
  });

  test("build() НЕ содержит @startuml", () => {
    const puml = new MindMapBuilder().build();
    expect(puml).not.toContain("@startuml");
  });
});

// ── GanttBuilder ───────────────────────────────────────────────────────────────

describe("GanttBuilder", () => {
  test("build() использует @startgantt / @endgantt", () => {
    const puml = new GanttBuilder().task("T1", 5).build();
    expect(puml).toMatch(/^@startgantt/);
    expect(puml).toMatch(/@endgantt$/);
  });

  test("task() без startDay", () => {
    const puml = new GanttBuilder().task("Разработка", 10).build();
    expect(puml).toContain("[Разработка] lasts 10 days");
  });

  test("after() связывает задачи последовательно", () => {
    const puml = new GanttBuilder().after("T2", "T1", 3).build();
    expect(puml).toContain("[T2] starts at [T1]'s end and lasts 3 days");
  });

  test("milestone() привязывается к концу задачи", () => {
    const puml = new GanttBuilder().milestone("Релиз", "T2").build();
    expect(puml).toContain("[Релиз] happens at [T2]'s end");
  });

  test("section() добавляет раздел", () => {
    const puml = new GanttBuilder().section("Фаза 1").build();
    expect(puml).toContain("-- Фаза 1 --");
  });

  test("startDate() задаёт дату начала", () => {
    const puml = new GanttBuilder().startDate("2026-01-01").build();
    expect(puml).toContain("Project starts 2026-01-01");
  });

  test("build() НЕ содержит @startuml", () => {
    const puml = new GanttBuilder().build();
    expect(puml).not.toContain("@startuml");
  });
});

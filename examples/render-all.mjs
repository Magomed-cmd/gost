import { createPlantUmlGost, ER } from "../dist/plantuml-gost.js";
import fs from "fs";
import path from "path";

const p = createPlantUmlGost({ dpi: 150 });
const OUT = path.resolve("examples/out");
fs.mkdirSync(OUT, { recursive: true });

async function save(name, puml) {
  const buf = await p.render(puml);
  fs.writeFileSync(path.join(OUT, `${name}.png`), buf);
  console.log(`✓ ${name}.png`);
}

// 1. Activity
const activity = p.activityDiagram()
  .start()
  .action("Ввести логин и пароль")
  .if("Данные корректны",
    (b) => b.action("Открыть главную страницу").action("Записать сессию"),
    (b) => b.action("Показать ошибку").action("Счётчик попыток +1"),
    "да", "нет")
  .action("Завершение")
  .stop();

// 2. Component
const component = p.componentDiagram()
  .package("Клиент", b => b.component("Browser", "BR"))
  .package("Сервер", b => {
    b.component("Express API", "API");
    b.component("Auth Service", "AUTH");
    b.database("PostgreSQL", "DB");
  })
  .cloud("Внешние сервисы", "EXT")
  .blank()
  .arrow("BR", "API", "HTTP/REST")
  .arrow("API", "AUTH", "проверка токена")
  .arrow("API", "DB", "SQL")
  .arrow("AUTH", "EXT", "OAuth2");

// 3. Deployment
const deployment = p.deploymentDiagram()
  .node("Клиент", "CLIENT", b => b.component("React App", "REACT"))
  .node("Сервер приложений", "APP_SRV", b => {
    b.component("Node.js API", "API");
    b.artifact("app.jar");
  })
  .node("База данных", "DB_SRV", b => b.database("PostgreSQL", "PG"))
  .cloud("CDN", b => b.artifact("static files"))
  .blank()
  .arrow("CLIENT", "APP_SRV", "HTTPS :443")
  .arrow("APP_SRV", "DB_SRV", "TCP :5432");

// 4. Object — простые имена без колонок
const object = p.objectDiagram()
  .object("tram1 : Tram", "tram1", { id: 1, state: '"ON_LINE"', route: 3 })
  .object("driver1 : Driver", "driver1", { id: 42, name: '"Иванов"' })
  .object("stop1 : TramStop", "stop1", { id: 101, name: '"Площадь"' })
  .blank()
  .link("tram1", "driver1", "управляет")
  .link("tram1", "stop1", "следующая");

// 5. ER
const er = p.erDiagram()
  .entity("User", ["id : int <<PK>>"], ["name : varchar(100)", "email : varchar(100)"])
  .entity("Order", ["id : int <<PK>>"], ["user_id : int <<FK>>", "total : decimal", "status : varchar(20)"])
  .entity("Product", ["id : int <<PK>>"], ["name : varchar(200)", "price : decimal"])
  .entity("OrderItem", ["order_id : int <<FK>>", "product_id : int <<FK>>"], ["quantity : int", "price : decimal"])
  .blank()
  .relation("User",    ER.ONE_TO_MANY, "Order",     "делает")
  .relation("Order",   ER.ONE_TO_MANY_REQ, "OrderItem", "содержит")
  .relation("Product", ER.ONE_TO_MANY, "OrderItem", "входит в");

// 6. Timing — без highlight (не поддерживается)
const timing = p.timingDiagram()
  .concise("Клиент", "C")
  .concise("Сервер", "S")
  .blank()
  .at(0).state("C", "Idle").state("S", "Listening")
  .at(50).state("C", "Waiting").arrow("C", "S", "HTTP GET /api")
  .at(150).state("S", "Processing")
  .at(250).state("S", "Listening").state("C", "Active").arrow("S", "C", "200 OK");

// 7. MindMap
const mindmap = p.mindMap()
  .node(1, "Паттерны GoF")
  .node(2, "Порождающие")
  .node(3, "Singleton").node(3, "Factory Method").node(3, "Abstract Factory").node(3, "Builder")
  .node(2, "Структурные")
  .node(3, "Adapter").node(3, "Decorator").node(3, "Facade")
  .node(2, "Поведенческие", "left")
  .node(3, "Observer", "left").node(3, "Strategy", "left").node(3, "State", "left").node(3, "Command", "left");

// 8. Gantt — только after() для цепочек
const gantt = p.ganttChart()
  .startDate("2026-09-01")
  .section("Подготовка")
  .task("Анализ требований", 5)
  .after("Проектирование БД", "Анализ требований", 3)
  .section("Разработка")
  .after("Backend API", "Проектирование БД", 14)
  .after("Frontend", "Backend API", 10)
  .after("Интеграция", "Frontend", 5)
  .section("Завершение")
  .after("Тестирование", "Интеграция", 7)
  .milestone("Релиз v1.0", "Тестирование");

await save("1_activity",   activity.build());
await save("2_component",  component.build());
await save("3_deployment", deployment.build());
await save("4_object",     object.build());
await save("5_er",         er.build());
await save("6_timing",     timing.build());
await save("7_mindmap",    mindmap.build());
await save("8_gantt",      gantt.build());
console.log("\nГотово → examples/out/");

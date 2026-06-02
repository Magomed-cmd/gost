import { renderDiagram, RenderDiagramOpts, SkinparamOpts } from "./plantuml-render";
import { autoImageSize } from "./plantuml-size";

// ── Типы ──────────────────────────────────────────────────────────────────────

export interface DiagramBuildOpts {
  /** Дополнительные skinparam поверх стандартных */
  skinparams?: Record<string, string>;
}

export interface PlantUmlGostOptions {
  /** DPI для растеризации. По умолчанию: 150 */
  dpi?: number;
}

export interface PlantUmlGostInstance {
  useCaseDiagram():    UseCaseBuilder;
  classDiagram():      ClassDiagramBuilder;
  sequenceDiagram():   SequenceBuilder;
  stateDiagram():      StateBuilder;
  activityDiagram():   ActivityBuilder;
  componentDiagram():  ComponentBuilder;
  deploymentDiagram(): DeploymentBuilder;
  objectDiagram():     ObjectBuilder;
  erDiagram():         ERBuilder;
  timingDiagram():     TimingBuilder;
  mindMap():           MindMapBuilder;
  ganttChart():        GanttBuilder;
  /** Рендерит puml-строку, возвращает PNG Buffer */
  render(puml: string, opts?: RenderDiagramOpts & SkinparamOpts): Promise<Buffer>;
  /** Авторазмер PNG под ширину контента */
  autoImageSize(buffer: Buffer, maxWidth?: number): { width: number; height: number };
}

// ── Базовый строитель ─────────────────────────────────────────────────────────

export class DiagramBuilder {
  protected readonly lines: string[] = [];

  /** Добавить произвольную строку PlantUML (escape hatch) */
  add(line: string): this {
    this.lines.push(line);
    return this;
  }

  /** Добавить пустую строку для визуального разделения */
  blank(): this {
    return this.add("");
  }

  /** Собрать итоговую строку @startuml...@enduml */
  build(): string {
    return `@startuml\n${this.lines.join("\n")}\n@enduml`;
  }
}

// ── Use Case builder ──────────────────────────────────────────────────────────

export class UseCaseBuilder extends DiagramBuilder {
  /** Объявить актора */
  actor(name: string, alias?: string): this {
    return this.add(alias ? `actor "${name}" as ${alias}` : `actor "${name}"`);
  }

  /** Объявить прецедент */
  usecase(name: string, alias?: string): this {
    return this.add(alias ? `usecase "${name}" as ${alias}` : `usecase "${name}"`);
  }

  /** Стрелка связи от from к to */
  arrow(from: string, to: string, label?: string): this {
    return this.add(label ? `${from} --> ${to} : ${label}` : `${from} --> ${to}`);
  }

  /** Отношение <<extend>> */
  extend(base: string, extension: string): this {
    return this.add(`${extension} .> ${base} : <<extend>>`);
  }

  /** Отношение <<include>> */
  include(base: string, included: string): this {
    return this.add(`${base} .> ${included} : <<include>>`);
  }

  /** Сгруппировать прецеденты в прямоугольник системы */
  package(name: string, fn: (b: UseCaseBuilder) => void): this {
    this.add(`rectangle "${name}" {`);
    fn(this);
    return this.add("}");
  }
}

// ── Class diagram builder ─────────────────────────────────────────────────────

export class ClassDiagramBuilder extends DiagramBuilder {
  /** build всегда добавляет hide circle (ГОСТ стиль) */
  override build(): string {
    return `@startuml\nhide circle\n${this.lines.join("\n")}\n@enduml`;
  }

  /** Объявить класс с необязательными членами */
  class(name: string, members: string[] = []): this {
    if (members.length === 0) return this.add(`class ${name}`);
    this.add(`class ${name} {`);
    members.forEach((m) => this.add(`  ${m}`));
    return this.add("}");
  }

  /** Объявить абстрактный класс */
  abstract(name: string, members: string[] = []): this {
    if (members.length === 0) return this.add(`abstract class ${name}`);
    this.add(`abstract class ${name} {`);
    members.forEach((m) => this.add(`  ${m}`));
    return this.add("}");
  }

  /** Объявить интерфейс */
  interface(name: string, members: string[] = []): this {
    if (members.length === 0) return this.add(`interface ${name}`);
    this.add(`interface ${name} {`);
    members.forEach((m) => this.add(`  ${m}`));
    return this.add("}");
  }

  /** Объявить перечисление */
  enum(name: string, values: string[]): this {
    this.add(`enum ${name} {`);
    values.forEach((v) => this.add(`  ${v}`));
    return this.add("}");
  }

  /** Наследование: child --|> parent */
  extends(child: string, parent: string): this {
    return this.add(`${child} --|> ${parent}`);
  }

  /** Реализация интерфейса: cls ..|> iface */
  implements(cls: string, iface: string): this {
    return this.add(`${cls} ..|> ${iface}`);
  }

  /** Ассоциация: from --> to */
  association(from: string, to: string, label?: string): this {
    return this.add(label ? `${from} --> ${to} : ${label}` : `${from} --> ${to}`);
  }

  /** Зависимость: from ..> to */
  dependency(from: string, to: string, label?: string): this {
    return this.add(label ? `${from} ..> ${to} : ${label}` : `${from} ..> ${to}`);
  }

  /** Композиция: from *-- to */
  composition(from: string, to: string, label?: string): this {
    return this.add(label ? `${from} *-- ${to} : ${label}` : `${from} *-- ${to}`);
  }

  /** Агрегация: from o-- to */
  aggregation(from: string, to: string, label?: string): this {
    return this.add(label ? `${from} o-- ${to} : ${label}` : `${from} o-- ${to}`);
  }

  /** Сгруппировать классы в пакет */
  package(name: string, fn: (b: ClassDiagramBuilder) => void): this {
    this.add(`package ${name} {`);
    fn(this);
    return this.add("}");
  }
}

// ── Sequence builder ──────────────────────────────────────────────────────────

export class SequenceBuilder extends DiagramBuilder {
  /** Объявить участника */
  participant(name: string, alias?: string): this {
    return this.add(alias ? `participant "${name}" as ${alias}` : `participant "${name}"`);
  }

  /** Объявить актора */
  actor(name: string, alias?: string): this {
    return this.add(alias ? `actor "${name}" as ${alias}` : `actor "${name}"`);
  }

  /** Синхронная стрелка from -> to */
  arrow(from: string, to: string, label: string): this {
    return this.add(`${from} -> ${to} : ${label}`);
  }

  /** Пунктирная возвратная стрелка from --> to */
  return(from: string, to: string, label: string): this {
    return this.add(`${from} --> ${to} : ${label}`);
  }

  /** Активировать участника */
  activate(name: string): this {
    return this.add(`activate ${name}`);
  }

  /** Деактивировать участника */
  deactivate(name: string): this {
    return this.add(`deactivate ${name}`);
  }

  /** Заметка рядом с участником */
  note(text: string, side: "left" | "right", target: string): this {
    return this.add(`note ${side} of ${target} : ${text}`);
  }

  /** Разделитель с подписью */
  divider(text?: string): this {
    return this.add(text ? `== ${text} ==` : "====");
  }

  /** Группа с меткой */
  group(label: string, fn: (b: SequenceBuilder) => void): this {
    this.add(`group ${label}`);
    fn(this);
    return this.add("end");
  }
}

// ── State builder ─────────────────────────────────────────────────────────────

export class StateBuilder extends DiagramBuilder {
  /** Объявить состояние */
  state(name: string, label?: string): this {
    return this.add(label ? `state "${label}" as ${name}` : `state ${name}`);
  }

  /** Переход между состояниями с явным направлением */
  transition(from: string, to: string, label?: string, direction?: "down" | "up" | "left" | "right"): this {
    const arrow = direction ? `-${direction}->` : "-->";
    return this.add(label ? `${from} ${arrow} ${to} : ${label}` : `${from} ${arrow} ${to}`);
  }

  /** Начальное состояние */
  initial(to: string): this {
    return this.add(`[*] --> ${to}`);
  }

  /** Конечное состояние */
  final(from: string): this {
    return this.add(`${from} --> [*]`);
  }
}

// ── Activity builder ──────────────────────────────────────────────────────────

export class ActivityBuilder extends DiagramBuilder {
  /** Начало диаграммы */
  start(): this { return this.add("start"); }

  /** Нормальное завершение (stop-узел) */
  stop(): this { return this.add("stop"); }

  /** Конечный узел (end-узел c завершением потока) */
  end(): this { return this.add("end"); }

  /** Действие — прямоугольник с текстом */
  action(text: string): this { return this.add(`:${text};`); }

  /** Заметка рядом с текущим действием */
  note(text: string, side: "right" | "left" = "right"): this {
    return this.add(`note ${side}: ${text}`);
  }

  /** Условное ветвление. noFn и метки опциональны */
  if(
    condition: string,
    yesFn: (b: ActivityBuilder) => void,
    noFn?: (b: ActivityBuilder) => void,
    yesLabel = "yes",
    noLabel = "no",
  ): this {
    this.add(`if (${condition}?) then (${yesLabel})`);
    yesFn(this);
    if (noFn) {
      this.add(`else (${noLabel})`);
      noFn(this);
    }
    return this.add("endif");
  }

  /** Цикл while */
  while(condition: string, fn: (b: ActivityBuilder) => void, exitLabel?: string): this {
    this.add(`while (${condition}?)`);
    fn(this);
    return this.add(exitLabel ? `endwhile (${exitLabel})` : "endwhile");
  }

  /** Параллельное ветвление: каждый элемент массива — одна ветвь */
  fork(...branches: Array<(b: ActivityBuilder) => void>): this {
    branches.forEach((branch, i) => {
      this.add(i === 0 ? "fork" : "fork again");
      branch(this);
    });
    return this.add("end fork");
  }

  /** Swimlane / раздел */
  partition(name: string, fn: (b: ActivityBuilder) => void): this {
    this.add(`partition "${name}" {`);
    fn(this);
    return this.add("}");
  }
}

// ── Component builder ─────────────────────────────────────────────────────────

export class ComponentBuilder extends DiagramBuilder {
  /** Компонент */
  component(name: string, alias?: string): this {
    return this.add(alias ? `component [${name}] as ${alias}` : `component [${name}]`);
  }

  /** Интерфейс */
  interface(name: string, alias?: string): this {
    return this.add(alias ? `interface "${name}" as ${alias}` : `interface "${name}"`);
  }

  /** База данных */
  database(name: string, alias?: string): this {
    return this.add(alias ? `database "${name}" as ${alias}` : `database "${name}"`);
  }

  /** Облако / внешняя система */
  cloud(name: string, alias?: string): this {
    return this.add(alias ? `cloud "${name}" as ${alias}` : `cloud "${name}"`);
  }

  /** Направленная зависимость */
  arrow(from: string, to: string, label?: string): this {
    return this.add(label ? `${from} --> ${to} : ${label}` : `${from} --> ${to}`);
  }

  /** Ненаправленная линия связи */
  link(from: string, to: string, label?: string): this {
    return this.add(label ? `${from} -- ${to} : ${label}` : `${from} -- ${to}`);
  }

  /** Сгруппировать в пакет / подсистему */
  package(name: string, fn: (b: ComponentBuilder) => void): this {
    this.add(`package "${name}" {`);
    fn(this);
    return this.add("}");
  }
}

// ── Deployment builder ────────────────────────────────────────────────────────

export class DeploymentBuilder extends DiagramBuilder {
  /** Физический узел (сервер, устройство) с опциональным вложением */
  node(name: string, alias?: string, fn?: (b: DeploymentBuilder) => void): this {
    if (fn) {
      this.add(alias ? `node "${name}" as ${alias} {` : `node "${name}" {`);
      fn(this);
      return this.add("}");
    }
    return this.add(alias ? `node "${name}" as ${alias}` : `node "${name}"`);
  }

  /** Компонент внутри узла */
  component(name: string, alias?: string): this {
    return this.add(alias ? `component [${name}] as ${alias}` : `component [${name}]`);
  }

  /** База данных */
  database(name: string, alias?: string): this {
    return this.add(alias ? `database "${name}" as ${alias}` : `database "${name}"`);
  }

  /** Артефакт (jar, war, исполняемый файл) */
  artifact(name: string, alias?: string): this {
    return this.add(alias ? `artifact "${name}" as ${alias}` : `artifact "${name}"`);
  }

  /** Облачная зона с опциональным вложением */
  cloud(name: string, fn?: (b: DeploymentBuilder) => void): this {
    if (fn) {
      this.add(`cloud "${name}" {`);
      fn(this);
      return this.add("}");
    }
    return this.add(`cloud "${name}"`);
  }

  /** Направленное соединение */
  arrow(from: string, to: string, label?: string): this {
    return this.add(label ? `${from} --> ${to} : ${label}` : `${from} --> ${to}`);
  }

  /** Физическое соединение без стрелки */
  link(from: string, to: string, label?: string): this {
    return this.add(label ? `${from} -- ${to} : ${label}` : `${from} -- ${to}`);
  }
}

// ── Object builder ────────────────────────────────────────────────────────────

export class ObjectBuilder extends DiagramBuilder {
  /**
   * Объект-экземпляр.
   * @param name   - идентификатор (используется в link/arrow). Если содержит
   *                 пробелы или двоеточие — автоматически оборачивается в кавычки,
   *                 а для связей нужно передать alias.
   * @param alias  - короткий идентификатор для link/arrow когда name содержит спецсимволы
   * @param fields - поля в формате { key: value }
   */
  object(name: string, alias?: string, fields: Record<string, string | number> = {}): this {
    const entries = Object.entries(fields);
    const needsQuotes = /[^a-zA-Z0-9_]/.test(name);
    const decl = needsQuotes
      ? `object "${name}"${alias ? ` as ${alias}` : ""}`
      : `object ${name}`;
    if (entries.length === 0) return this.add(decl);
    this.add(`${decl} {`);
    entries.forEach(([k, v]) => this.add(`  ${k} = ${v}`));
    return this.add("}");
  }

  /** Связь между объектами */
  link(from: string, to: string, label?: string): this {
    return this.add(label ? `${from} --> ${to} : ${label}` : `${from} --> ${to}`);
  }
}

// ── ER builder ────────────────────────────────────────────────────────────────

/**
 * Готовые строки связей для ER-диаграмм PlantUML.
 * Используй как второй аргумент ERBuilder.relation().
 *
 * В PlantUML нотация кардинальности зеркальна:
 *   левая сторона (--): `||`, `|o`, `}|`, `}o`
 *   правая сторона (--): `||`, `o|`, `|{`, `o{`
 * Константы ниже — готовые комбинации, которые всегда корректны.
 *
 * @example er.relation("User", ER.ONE_TO_MANY, "Order", "делает")
 */
export const ER = {
  /** || -- ||  один к одному */
  ONE_TO_ONE:        "||--||",
  /** || -- o{  один к нулю-или-многим (самый частый) */
  ONE_TO_MANY:       "||--o{",
  /** || -- |{  один к одному-или-многим */
  ONE_TO_MANY_REQ:   "||--|{",
  /** |o -- o{  ноль-или-один к нулю-или-многим */
  OPT_TO_MANY:       "|o--o{",
  /** }o -- o{  многие ко многим */
  MANY_TO_MANY:      "}o--o{",
} as const;

export class ERBuilder extends DiagramBuilder {
  /**
   * Объявить сущность.
   * @param pkFields - поля первичного ключа (выделяются * и отделяются чертой)
   * @param fields   - обычные поля
   */
  entity(name: string, pkFields: string[], fields: string[] = []): this {
    this.add(`entity "${name}" {`);
    pkFields.forEach((f) => this.add(`  * ${f}`));
    if (pkFields.length > 0 && fields.length > 0) this.add("  --");
    fields.forEach((f) => this.add(`  ${f}`));
    return this.add("}");
  }

  /**
   * Связь между сущностями.
   * @param rel - строка кардинальности, используй константы ER.* или свою строку
   * @example er.relation("User", ER.ONE_TO_MANY, "Order", "делает")
   */
  relation(from: string, rel: string, to: string, label?: string): this {
    return this.add(label
      ? `${from} ${rel} ${to} : ${label}`
      : `${from} ${rel} ${to}`
    );
  }
}

// ── Timing builder ────────────────────────────────────────────────────────────

export class TimingBuilder extends DiagramBuilder {
  /** Участник типа concise — переключается между дискретными состояниями */
  concise(name: string, alias?: string): this {
    return this.add(alias ? `concise "${name}" as ${alias}` : `concise "${name}"`);
  }

  /** Участник типа robust — аналоговый сигнал с длительностью переходов */
  robust(name: string, alias?: string): this {
    return this.add(alias ? `robust "${name}" as ${alias}` : `robust "${name}"`);
  }

  /** Тактовый сигнал с периодом */
  clock(name: string, period: number): this {
    return this.add(`clock "${name}" with period ${period}`);
  }

  /** Временна́я точка — все последующие state() привязаны к этому моменту */
  at(time: number | string): this {
    return this.add(`@${time}`);
  }

  /** Состояние участника в текущей временно́й точке */
  state(participant: string, value: string): this {
    return this.add(`${participant} is ${value}`);
  }

  /** Сообщение между участниками */
  arrow(from: string, to: string, label: string): this {
    return this.add(`${from} -> ${to} : ${label}`);
  }

  /**
   * Выделить временной отрезок цветом (поддерживается в PlantUML 1.2021+).
   * Синтаксис: highlight start to end #Color : label
   */
  highlight(fromTime: number, toTime: number, color = "#LightYellow", label?: string): this {
    return this.add(label
      ? `highlight ${fromTime} to ${toTime} ${color} : ${label}`
      : `highlight ${fromTime} to ${toTime} ${color}`
    );
  }
}

// ── MindMap builder ───────────────────────────────────────────────────────────

export class MindMapBuilder extends DiagramBuilder {
  /** MindMap использует собственные теги вместо @startuml */
  override build(): string {
    return `@startmindmap\n${this.lines.join("\n")}\n@endmindmap`;
  }

  /**
   * Узел карты.
   * @param level - глубина: 1 = корень, 2 = ветвь, 3 = лист, ...
   * @param text  - текст узла
   * @param side  - "right" (по умолчанию) или "left" (ветви идут влево от корня)
   */
  node(level: number, text: string, side: "right" | "left" = "right"): this {
    const stars = "*".repeat(Math.max(1, level));
    return this.add(side === "left" ? `${stars}_ ${text}` : `${stars} ${text}`);
  }
}

// ── Gantt builder ─────────────────────────────────────────────────────────────

export class GanttBuilder extends DiagramBuilder {
  /** Gantt использует собственные теги вместо @startuml */
  override build(): string {
    return `@startgantt\n${this.lines.join("\n")}\n@endgantt`;
  }

  /** Дата начала проекта в формате YYYY-MM-DD */
  startDate(date: string): this {
    return this.add(`Project starts ${date}`);
  }

  /** Задача с длительностью в днях. Начинается сразу после предыдущей */
  task(name: string, duration: number): this {
    return this.add(`[${name}] lasts ${duration} days`);
  }

  /** Задача, стартующая сразу после окончания другой */
  after(name: string, afterName: string, duration: number): this {
    return this.add(`[${name}] starts at [${afterName}]'s end and lasts ${duration} days`);
  }

  /** Веха — нулевая длительность в момент окончания задачи */
  milestone(name: string, afterTask: string): this {
    return this.add(`[${name}] happens at [${afterTask}]'s end`);
  }

  /** Раздел / фаза проекта */
  section(name: string): this {
    return this.add(`-- ${name} --`);
  }
}

// ── Фабрика ───────────────────────────────────────────────────────────────────

export function createPlantUmlGost(opts: PlantUmlGostOptions = {}): PlantUmlGostInstance {
  const defaultDpi = opts.dpi ?? 150;

  return {
    useCaseDiagram:    () => new UseCaseBuilder(),
    classDiagram:      () => new ClassDiagramBuilder(),
    sequenceDiagram:   () => new SequenceBuilder(),
    stateDiagram:      () => new StateBuilder(),
    activityDiagram:   () => new ActivityBuilder(),
    componentDiagram:  () => new ComponentBuilder(),
    deploymentDiagram: () => new DeploymentBuilder(),
    objectDiagram:     () => new ObjectBuilder(),
    erDiagram:         () => new ERBuilder(),
    timingDiagram:     () => new TimingBuilder(),
    mindMap:           () => new MindMapBuilder(),
    ganttChart:        () => new GanttBuilder(),

    render: (puml, renderOpts = {}) =>
      renderDiagram(puml, { dpi: defaultDpi, ...renderOpts }),

    autoImageSize: (buffer, maxWidth) =>
      autoImageSize(buffer, maxWidth),
  };
}

export { autoImageSize } from "./plantuml-size";
export { renderDiagram, wrapWithSkin } from "./plantuml-render";

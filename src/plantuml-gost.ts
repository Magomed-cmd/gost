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
  useCaseDiagram(): UseCaseBuilder;
  classDiagram(): ClassDiagramBuilder;
  sequenceDiagram(): SequenceBuilder;
  stateDiagram(): StateBuilder;
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

// ── Фабрика ───────────────────────────────────────────────────────────────────

export function createPlantUmlGost(opts: PlantUmlGostOptions = {}): PlantUmlGostInstance {
  const defaultDpi = opts.dpi ?? 150;

  return {
    useCaseDiagram:  () => new UseCaseBuilder(),
    classDiagram:    () => new ClassDiagramBuilder(),
    sequenceDiagram: () => new SequenceBuilder(),
    stateDiagram:    () => new StateBuilder(),

    render: (puml, renderOpts = {}) =>
      renderDiagram(puml, { dpi: defaultDpi, ...renderOpts }),

    autoImageSize: (buffer, maxWidth) =>
      autoImageSize(buffer, maxWidth),
  };
}

export { autoImageSize } from "./plantuml-size";
export { renderDiagram, wrapWithSkin } from "./plantuml-render";

"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.wrapWithSkin = exports.renderDiagram = exports.autoImageSize = exports.StateBuilder = exports.SequenceBuilder = exports.ClassDiagramBuilder = exports.UseCaseBuilder = exports.DiagramBuilder = void 0;
exports.createPlantUmlGost = createPlantUmlGost;
const plantuml_render_1 = require("./plantuml-render");
const plantuml_size_1 = require("./plantuml-size");
// ── Базовый строитель ─────────────────────────────────────────────────────────
class DiagramBuilder {
    constructor() {
        this.lines = [];
    }
    /** Добавить произвольную строку PlantUML (escape hatch) */
    add(line) {
        this.lines.push(line);
        return this;
    }
    /** Добавить пустую строку для визуального разделения */
    blank() {
        return this.add("");
    }
    /** Собрать итоговую строку @startuml...@enduml */
    build() {
        return `@startuml\n${this.lines.join("\n")}\n@enduml`;
    }
}
exports.DiagramBuilder = DiagramBuilder;
// ── Use Case builder ──────────────────────────────────────────────────────────
class UseCaseBuilder extends DiagramBuilder {
    /** Объявить актора */
    actor(name, alias) {
        return this.add(alias ? `actor "${name}" as ${alias}` : `actor "${name}"`);
    }
    /** Объявить прецедент */
    usecase(name, alias) {
        return this.add(alias ? `usecase "${name}" as ${alias}` : `usecase "${name}"`);
    }
    /** Стрелка связи от from к to */
    arrow(from, to, label) {
        return this.add(label ? `${from} --> ${to} : ${label}` : `${from} --> ${to}`);
    }
    /** Отношение <<extend>> */
    extend(base, extension) {
        return this.add(`${extension} .> ${base} : <<extend>>`);
    }
    /** Отношение <<include>> */
    include(base, included) {
        return this.add(`${base} .> ${included} : <<include>>`);
    }
    /** Сгруппировать прецеденты в прямоугольник системы */
    package(name, fn) {
        this.add(`rectangle "${name}" {`);
        fn(this);
        return this.add("}");
    }
}
exports.UseCaseBuilder = UseCaseBuilder;
// ── Class diagram builder ─────────────────────────────────────────────────────
class ClassDiagramBuilder extends DiagramBuilder {
    /** build всегда добавляет hide circle (ГОСТ стиль) */
    build() {
        return `@startuml\nhide circle\n${this.lines.join("\n")}\n@enduml`;
    }
    /** Объявить класс с необязательными членами */
    class(name, members = []) {
        if (members.length === 0)
            return this.add(`class ${name}`);
        this.add(`class ${name} {`);
        members.forEach((m) => this.add(`  ${m}`));
        return this.add("}");
    }
    /** Объявить абстрактный класс */
    abstract(name, members = []) {
        if (members.length === 0)
            return this.add(`abstract class ${name}`);
        this.add(`abstract class ${name} {`);
        members.forEach((m) => this.add(`  ${m}`));
        return this.add("}");
    }
    /** Объявить интерфейс */
    interface(name, members = []) {
        if (members.length === 0)
            return this.add(`interface ${name}`);
        this.add(`interface ${name} {`);
        members.forEach((m) => this.add(`  ${m}`));
        return this.add("}");
    }
    /** Объявить перечисление */
    enum(name, values) {
        this.add(`enum ${name} {`);
        values.forEach((v) => this.add(`  ${v}`));
        return this.add("}");
    }
    /** Наследование: child --|> parent */
    extends(child, parent) {
        return this.add(`${child} --|> ${parent}`);
    }
    /** Реализация интерфейса: cls ..|> iface */
    implements(cls, iface) {
        return this.add(`${cls} ..|> ${iface}`);
    }
    /** Ассоциация: from --> to */
    association(from, to, label) {
        return this.add(label ? `${from} --> ${to} : ${label}` : `${from} --> ${to}`);
    }
    /** Зависимость: from ..> to */
    dependency(from, to, label) {
        return this.add(label ? `${from} ..> ${to} : ${label}` : `${from} ..> ${to}`);
    }
    /** Композиция: from *-- to */
    composition(from, to, label) {
        return this.add(label ? `${from} *-- ${to} : ${label}` : `${from} *-- ${to}`);
    }
    /** Агрегация: from o-- to */
    aggregation(from, to, label) {
        return this.add(label ? `${from} o-- ${to} : ${label}` : `${from} o-- ${to}`);
    }
    /** Сгруппировать классы в пакет */
    package(name, fn) {
        this.add(`package ${name} {`);
        fn(this);
        return this.add("}");
    }
}
exports.ClassDiagramBuilder = ClassDiagramBuilder;
// ── Sequence builder ──────────────────────────────────────────────────────────
class SequenceBuilder extends DiagramBuilder {
    /** Объявить участника */
    participant(name, alias) {
        return this.add(alias ? `participant "${name}" as ${alias}` : `participant "${name}"`);
    }
    /** Объявить актора */
    actor(name, alias) {
        return this.add(alias ? `actor "${name}" as ${alias}` : `actor "${name}"`);
    }
    /** Синхронная стрелка from -> to */
    arrow(from, to, label) {
        return this.add(`${from} -> ${to} : ${label}`);
    }
    /** Пунктирная возвратная стрелка from --> to */
    return(from, to, label) {
        return this.add(`${from} --> ${to} : ${label}`);
    }
    /** Активировать участника */
    activate(name) {
        return this.add(`activate ${name}`);
    }
    /** Деактивировать участника */
    deactivate(name) {
        return this.add(`deactivate ${name}`);
    }
    /** Заметка рядом с участником */
    note(text, side, target) {
        return this.add(`note ${side} of ${target} : ${text}`);
    }
    /** Разделитель с подписью */
    divider(text) {
        return this.add(text ? `== ${text} ==` : "====");
    }
    /** Группа с меткой */
    group(label, fn) {
        this.add(`group ${label}`);
        fn(this);
        return this.add("end");
    }
}
exports.SequenceBuilder = SequenceBuilder;
// ── State builder ─────────────────────────────────────────────────────────────
class StateBuilder extends DiagramBuilder {
    /** Объявить состояние */
    state(name, label) {
        return this.add(label ? `state "${label}" as ${name}` : `state ${name}`);
    }
    /** Переход между состояниями с явным направлением */
    transition(from, to, label, direction) {
        const arrow = direction ? `-${direction}->` : "-->";
        return this.add(label ? `${from} ${arrow} ${to} : ${label}` : `${from} ${arrow} ${to}`);
    }
    /** Начальное состояние */
    initial(to) {
        return this.add(`[*] --> ${to}`);
    }
    /** Конечное состояние */
    final(from) {
        return this.add(`${from} --> [*]`);
    }
}
exports.StateBuilder = StateBuilder;
// ── Фабрика ───────────────────────────────────────────────────────────────────
function createPlantUmlGost(opts = {}) {
    const defaultDpi = opts.dpi ?? 150;
    return {
        useCaseDiagram: () => new UseCaseBuilder(),
        classDiagram: () => new ClassDiagramBuilder(),
        sequenceDiagram: () => new SequenceBuilder(),
        stateDiagram: () => new StateBuilder(),
        render: (puml, renderOpts = {}) => (0, plantuml_render_1.renderDiagram)(puml, { dpi: defaultDpi, ...renderOpts }),
        autoImageSize: (buffer, maxWidth) => (0, plantuml_size_1.autoImageSize)(buffer, maxWidth),
    };
}
var plantuml_size_2 = require("./plantuml-size");
Object.defineProperty(exports, "autoImageSize", { enumerable: true, get: function () { return plantuml_size_2.autoImageSize; } });
var plantuml_render_2 = require("./plantuml-render");
Object.defineProperty(exports, "renderDiagram", { enumerable: true, get: function () { return plantuml_render_2.renderDiagram; } });
Object.defineProperty(exports, "wrapWithSkin", { enumerable: true, get: function () { return plantuml_render_2.wrapWithSkin; } });
//# sourceMappingURL=plantuml-gost.js.map
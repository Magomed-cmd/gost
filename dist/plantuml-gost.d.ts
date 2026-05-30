import { RenderDiagramOpts, SkinparamOpts } from "./plantuml-render";
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
    autoImageSize(buffer: Buffer, maxWidth?: number): {
        width: number;
        height: number;
    };
}
export declare class DiagramBuilder {
    protected readonly lines: string[];
    /** Добавить произвольную строку PlantUML (escape hatch) */
    add(line: string): this;
    /** Добавить пустую строку для визуального разделения */
    blank(): this;
    /** Собрать итоговую строку @startuml...@enduml */
    build(): string;
}
export declare class UseCaseBuilder extends DiagramBuilder {
    /** Объявить актора */
    actor(name: string, alias?: string): this;
    /** Объявить прецедент */
    usecase(name: string, alias?: string): this;
    /** Стрелка связи от from к to */
    arrow(from: string, to: string, label?: string): this;
    /** Отношение <<extend>> */
    extend(base: string, extension: string): this;
    /** Отношение <<include>> */
    include(base: string, included: string): this;
    /** Сгруппировать прецеденты в прямоугольник системы */
    package(name: string, fn: (b: UseCaseBuilder) => void): this;
}
export declare class ClassDiagramBuilder extends DiagramBuilder {
    /** build всегда добавляет hide circle (ГОСТ стиль) */
    build(): string;
    /** Объявить класс с необязательными членами */
    class(name: string, members?: string[]): this;
    /** Объявить абстрактный класс */
    abstract(name: string, members?: string[]): this;
    /** Объявить интерфейс */
    interface(name: string, members?: string[]): this;
    /** Объявить перечисление */
    enum(name: string, values: string[]): this;
    /** Наследование: child --|> parent */
    extends(child: string, parent: string): this;
    /** Реализация интерфейса: cls ..|> iface */
    implements(cls: string, iface: string): this;
    /** Ассоциация: from --> to */
    association(from: string, to: string, label?: string): this;
    /** Зависимость: from ..> to */
    dependency(from: string, to: string, label?: string): this;
    /** Композиция: from *-- to */
    composition(from: string, to: string, label?: string): this;
    /** Агрегация: from o-- to */
    aggregation(from: string, to: string, label?: string): this;
    /** Сгруппировать классы в пакет */
    package(name: string, fn: (b: ClassDiagramBuilder) => void): this;
}
export declare class SequenceBuilder extends DiagramBuilder {
    /** Объявить участника */
    participant(name: string, alias?: string): this;
    /** Объявить актора */
    actor(name: string, alias?: string): this;
    /** Синхронная стрелка from -> to */
    arrow(from: string, to: string, label: string): this;
    /** Пунктирная возвратная стрелка from --> to */
    return(from: string, to: string, label: string): this;
    /** Активировать участника */
    activate(name: string): this;
    /** Деактивировать участника */
    deactivate(name: string): this;
    /** Заметка рядом с участником */
    note(text: string, side: "left" | "right", target: string): this;
    /** Разделитель с подписью */
    divider(text?: string): this;
    /** Группа с меткой */
    group(label: string, fn: (b: SequenceBuilder) => void): this;
}
export declare class StateBuilder extends DiagramBuilder {
    /** Объявить состояние */
    state(name: string, label?: string): this;
    /** Переход между состояниями с явным направлением */
    transition(from: string, to: string, label?: string, direction?: "down" | "up" | "left" | "right"): this;
    /** Начальное состояние */
    initial(to: string): this;
    /** Конечное состояние */
    final(from: string): this;
}
export declare function createPlantUmlGost(opts?: PlantUmlGostOptions): PlantUmlGostInstance;
export { autoImageSize } from "./plantuml-size";
export { renderDiagram, wrapWithSkin } from "./plantuml-render";
//# sourceMappingURL=plantuml-gost.d.ts.map
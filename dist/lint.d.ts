import { Paragraph, Table } from "docx";
export type GostElementType = "pageBreak" | "blank" | "h1" | "h2" | "h3" | "paragraph" | "centered" | "caption" | "tableCaption" | "placeholder" | "image" | "codeLine" | "table" | "unknown";
export type LintLevel = "warn" | "info";
export interface LintIssue {
    level: LintLevel;
    /** Индекс первого элемента, вызвавшего проблему */
    index: number;
    /** Машиночитаемый код проблемы */
    code: string;
    /** Человекочитаемое описание */
    message: string;
}
/**
 * Помечает элемент docx типом — используется внутри фабричных функций.
 * Возвращает тот же объект, не создаёт копию.
 */
export declare function tagElement<T extends object>(el: T, type: GostElementType): T;
/**
 * Анализирует массив children перед makeDocument и возвращает список проблем.
 * Не бросает исключений — только собирает предупреждения.
 */
export declare function lintChildren(children: Array<Paragraph | Table>): LintIssue[];
/**
 * Выводит результаты lint в консоль в читаемом виде.
 * Вызывать перед saveDocument.
 */
export declare function printLintResults(issues: LintIssue[]): void;
//# sourceMappingURL=lint.d.ts.map
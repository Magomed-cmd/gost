export interface RenderDiagramOpts {
    /** Формат вывода. По умолчанию: "png" */
    format?: "png" | "svg";
    /** DPI растеризации. По умолчанию: 150 */
    dpi?: number;
}
export interface SkinparamOpts {
    /** Переопределение или дополнение базовых skinparam */
    skinparams?: Record<string, string>;
}
/**
 * Оборачивает puml-источник в базовые skinparam.
 * Не нужно прописывать их в каждой диаграмме вручную.
 * opts.skinparams позволяет переопределить или добавить любой ключ.
 */
export declare function wrapWithSkin(source: string, dpi?: number, opts?: SkinparamOpts): string;
/**
 * Принимает строку PlantUML, возвращает PNG (или SVG) буфер.
 * Базовые skinparam добавляются автоматически через wrapWithSkin.
 */
export declare function renderDiagram(source: string, opts?: RenderDiagramOpts & SkinparamOpts): Promise<Buffer>;
//# sourceMappingURL=plantuml-render.d.ts.map
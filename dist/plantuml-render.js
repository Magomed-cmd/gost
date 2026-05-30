"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.wrapWithSkin = wrapWithSkin;
exports.renderDiagram = renderDiagram;
// eslint-disable-next-line @typescript-eslint/no-require-imports
const plantuml = require("node-plantuml");
// ── Базовый skinparam (ГОСТ + PlantUML defaults) ──────────────────────────────
const BASE_SKINPARAMS = {
    defaultFontName: '"Times New Roman"',
    defaultFontSize: "12",
    Shadowing: "false",
    monochrome: "false",
    nodesep: "60",
    ranksep: "50",
};
/**
 * Оборачивает puml-источник в базовые skinparam.
 * Не нужно прописывать их в каждой диаграмме вручную.
 * opts.skinparams позволяет переопределить или добавить любой ключ.
 */
function wrapWithSkin(source, dpi = 150, opts = {}) {
    const merged = { ...BASE_SKINPARAMS, dpi: String(dpi), ...opts.skinparams };
    const skin = Object.entries(merged)
        .map(([k, v]) => `skinparam ${k} ${v}`)
        .join("\n");
    return source.replace("@startuml", `@startuml\n${skin}`);
}
// ── Stream → Buffer ────────────────────────────────────────────────────────────
function streamToBuffer(stream) {
    return new Promise((resolve, reject) => {
        const chunks = [];
        stream.on("data", (chunk) => chunks.push(chunk));
        stream.on("end", () => resolve(Buffer.concat(chunks)));
        stream.on("error", reject);
    });
}
// ── Рендер ────────────────────────────────────────────────────────────────────
/**
 * Принимает строку PlantUML, возвращает PNG (или SVG) буфер.
 * Базовые skinparam добавляются автоматически через wrapWithSkin.
 */
async function renderDiagram(source, opts = {}) {
    const { format = "png", dpi = 150 } = opts;
    const wrapped = wrapWithSkin(source, dpi, { skinparams: opts.skinparams });
    const gen = plantuml.generate(wrapped, { format });
    return streamToBuffer(gen.out);
}
//# sourceMappingURL=plantuml-render.js.map
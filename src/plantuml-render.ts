import { Readable } from "stream";

// eslint-disable-next-line @typescript-eslint/no-require-imports
const plantuml = require("node-plantuml");

// ── Типы ──────────────────────────────────────────────────────────────────────

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

// ── Базовый skinparam (ГОСТ + PlantUML defaults) ──────────────────────────────

const BASE_SKINPARAMS: Record<string, string> = {
  defaultFontName: '"Times New Roman"',
  defaultFontSize: "12",
  Shadowing:       "false",
  monochrome:      "false",
  nodesep:         "60",
  ranksep:         "50",
};

/**
 * Оборачивает puml-источник в базовые skinparam.
 * Не нужно прописывать их в каждой диаграмме вручную.
 * opts.skinparams позволяет переопределить или добавить любой ключ.
 */
export function wrapWithSkin(source: string, dpi = 150, opts: SkinparamOpts = {}): string {
  const merged = { ...BASE_SKINPARAMS, dpi: String(dpi), ...opts.skinparams };
  const skin = Object.entries(merged)
    .map(([k, v]) => `skinparam ${k} ${v}`)
    .join("\n");
  return source.replace("@startuml", `@startuml\n${skin}`);
}

// ── Stream → Buffer ────────────────────────────────────────────────────────────

function streamToBuffer(stream: Readable): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    stream.on("data", (chunk: Buffer) => chunks.push(chunk));
    stream.on("end",  () => resolve(Buffer.concat(chunks)));
    stream.on("error", reject);
  });
}

// ── Рендер ────────────────────────────────────────────────────────────────────

/**
 * Принимает строку PlantUML, возвращает PNG (или SVG) буфер.
 * Базовые skinparam добавляются автоматически через wrapWithSkin.
 */
export async function renderDiagram(
  source: string,
  opts: RenderDiagramOpts & SkinparamOpts = {}
): Promise<Buffer> {
  const { format = "png", dpi = 150 } = opts;
  const wrapped = wrapWithSkin(source, dpi, { skinparams: opts.skinparams });
  const gen = plantuml.generate(wrapped, { format });
  return streamToBuffer(gen.out);
}

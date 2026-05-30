"use strict";
// Автоматическое определение размеров PNG и масштабирование под ширину контента.
// Максимальная ширина контента ГОСТ A4 с полями библиотеки = 624px (9355 DXA).
Object.defineProperty(exports, "__esModule", { value: true });
exports.autoImageSize = autoImageSize;
/**
 * Читает реальные размеры PNG из буфера (согласно PNG spec: bytes 16-20 = width, 20-24 = height)
 * и масштабирует пропорционально до maxWidth.
 */
function autoImageSize(buffer, maxWidth = 624) {
    const w = buffer.readUInt32BE(16);
    const h = buffer.readUInt32BE(20);
    const scale = maxWidth / w;
    return {
        width: maxWidth,
        height: Math.round(h * scale),
    };
}
//# sourceMappingURL=plantuml-size.js.map
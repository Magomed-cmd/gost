// Автоматическое определение размеров PNG и масштабирование под ширину контента.
// Максимальная ширина контента ГОСТ A4 с полями библиотеки = 624px (9355 DXA).

export interface ImageSize {
  width: number;
  height: number;
}

/**
 * Читает реальные размеры PNG из буфера (согласно PNG spec: bytes 16-20 = width, 20-24 = height)
 * и масштабирует пропорционально до maxWidth.
 */
export function autoImageSize(buffer: Buffer, maxWidth = 624): ImageSize {
  const w = buffer.readUInt32BE(16);
  const h = buffer.readUInt32BE(20);
  const scale = maxWidth / w;
  return {
    width: maxWidth,
    height: Math.round(h * scale),
  };
}

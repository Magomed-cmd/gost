export interface ImageSize {
    width: number;
    height: number;
}
/**
 * Читает реальные размеры PNG из буфера (согласно PNG spec: bytes 16-20 = width, 20-24 = height)
 * и масштабирует пропорционально до maxWidth.
 */
export declare function autoImageSize(buffer: Buffer, maxWidth?: number): ImageSize;
//# sourceMappingURL=plantuml-size.d.ts.map
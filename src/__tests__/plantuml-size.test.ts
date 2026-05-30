import { autoImageSize } from "../plantuml-size";

// Минимальный PNG-буфер с нужными размерами в нужных байтах (PNG spec: bytes 16-19 = width, 20-23 = height)
function fakePng(width: number, height: number): Buffer {
  const buf = Buffer.alloc(24);
  buf.writeUInt32BE(width, 16);
  buf.writeUInt32BE(height, 20);
  return buf;
}

describe("autoImageSize", () => {
  test("масштабирует до maxWidth по умолчанию (624)", () => {
    const buf = fakePng(800, 600);
    const { width, height } = autoImageSize(buf);
    expect(width).toBe(624);
    expect(height).toBe(Math.round(600 * (624 / 800)));
  });

  test("масштабирует до переданного maxWidth", () => {
    const buf = fakePng(1000, 500);
    const { width, height } = autoImageSize(buf, 500);
    expect(width).toBe(500);
    expect(height).toBe(250);
  });

  test("пропорции сохраняются при масштабировании", () => {
    const buf = fakePng(400, 200); // соотношение 2:1
    const { width, height } = autoImageSize(buf, 200);
    expect(width / height).toBeCloseTo(2, 1);
  });

  test("если изображение уже равно maxWidth — размеры не меняются", () => {
    const buf = fakePng(624, 400);
    const { width, height } = autoImageSize(buf, 624);
    expect(width).toBe(624);
    expect(height).toBe(400);
  });
});

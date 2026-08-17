import type { PptxSvgBox } from './render-svg-values';

export interface PptxSvgImageCrop {
  height: number;
  width: number;
  x: number;
  y: number;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Object(value) === value && !Array.isArray(value);
}

function cropValue(value: unknown): number | null {
  if (value === undefined) return 0;
  return typeof value === 'number' &&
    Number.isFinite(value) &&
    value >= -100 &&
    value <= 100
    ? value
    : null;
}

export function svgImageCrop(
  value: unknown,
  box: PptxSvgBox,
): PptxSvgImageCrop | null {
  if (!isRecord(value)) return null;
  const left = cropValue(value.l);
  const right = cropValue(value.r);
  const top = cropValue(value.t);
  const bottom = cropValue(value.b);
  if (left === null || right === null || top === null || bottom === null) {
    return null;
  }
  const horizontal = 1 - (left + right) / 100;
  const vertical = 1 - (top + bottom) / 100;
  if (horizontal <= 0 || vertical <= 0) return null;
  const width = box.width / horizontal;
  const height = box.height / vertical;
  return {
    height,
    width,
    x: left === 0 ? 0 : -(left / 100) * width,
    y: top === 0 ? 0 : -(top / 100) * height,
  };
}

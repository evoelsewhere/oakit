import { decodedBase64ByteLength } from '../../common/binary/base64';
import { normalizeHexColor } from '../../common/text/css';

export interface PptxSvgBox {
  height: number;
  left: number;
  top: number;
  width: number;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Object(value) === value;
}

export function svgNumber(value: unknown): string | null {
  if (typeof value !== 'number' || !Number.isFinite(value)) return null;
  const rounded = Math.round(value * 10_000) / 10_000;
  return String(rounded);
}

export function svgColor(value: unknown): string | null {
  return typeof value === 'string' ? normalizeHexColor(value) : null;
}

export function svgDashArray(value: unknown): string | null {
  if (typeof value !== 'string' || value.trim() === '') return null;
  const tokens = value.trim().split(/[ ,]+/);
  const numbers = tokens.map((token) => Number(token));
  if (numbers.some((number) => !Number.isFinite(number) || number < 0)) {
    return null;
  }
  return numbers.map((number) => svgNumber(number)).join(' ');
}

const RASTER_MEDIA_TYPES = new Set([
  'image/gif',
  'image/jpeg',
  'image/png',
  'image/webp',
]);

export function embeddedRasterDataUri(value: unknown): string | null {
  if (typeof value !== 'string' || !value.startsWith('data:')) return null;
  const separator = ';base64,';
  const separatorIndex = value.indexOf(separator);
  const mediaType = value.slice(5, separatorIndex).toLowerCase();
  if (!RASTER_MEDIA_TYPES.has(mediaType)) return null;
  const encoded = value.slice(separatorIndex + separator.length);
  try {
    if (decodedBase64ByteLength(encoded) === 0) return null;
  } catch {
    return null;
  }
  return `data:${mediaType};base64,${encoded}`;
}

export function svgBox(value: unknown): PptxSvgBox | null {
  if (!isRecord(value)) return null;
  const left = value.left;
  const top = value.top;
  const width = value.width;
  const height = value.height;
  if (
    typeof left !== 'number' ||
    !Number.isFinite(left) ||
    typeof top !== 'number' ||
    !Number.isFinite(top) ||
    typeof width !== 'number' ||
    !Number.isFinite(width) ||
    width <= 0 ||
    typeof height !== 'number' ||
    !Number.isFinite(height) ||
    height <= 0
  ) {
    return null;
  }
  return { height, left, top, width };
}

export function svgLineBox(value: unknown): PptxSvgBox | null {
  if (!isRecord(value)) return null;
  const left = value.left;
  const top = value.top;
  const width = value.width;
  const height = value.height;
  if (
    typeof left !== 'number' ||
    !Number.isFinite(left) ||
    typeof top !== 'number' ||
    !Number.isFinite(top) ||
    typeof width !== 'number' ||
    !Number.isFinite(width) ||
    width < 0 ||
    typeof height !== 'number' ||
    !Number.isFinite(height) ||
    height < 0 ||
    (width === 0 && height === 0)
  ) {
    return null;
  }
  return { height, left, top, width };
}

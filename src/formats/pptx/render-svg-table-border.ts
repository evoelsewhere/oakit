import { svgColor, svgNumber } from './render-svg-values';

function isRecord(value: unknown): value is Record<string, unknown> {
  return Object(value) === value && !Array.isArray(value);
}

function borderLine(
  border: unknown,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
): string {
  if (!isRecord(border)) return '';
  const color = svgColor(border.borderColor);
  const width = border.borderWidth;
  if (
    color === null ||
    typeof width !== 'number' ||
    !Number.isFinite(width) ||
    width <= 0
  ) {
    return '';
  }
  const dash =
    border.borderType === 'dashed'
      ? ' stroke-dasharray="4 3"'
      : border.borderType === 'dotted'
        ? ' stroke-dasharray="1 2"'
        : '';
  return `<line x1="${svgNumber(x1)}" y1="${svgNumber(y1)}" x2="${svgNumber(x2)}" y2="${svgNumber(y2)}" stroke="${color}" stroke-width="${svgNumber(width)}"${dash}/>`;
}

export function renderPptxSvgCellBorders(
  value: unknown,
  x: number,
  y: number,
  width: number,
  height: number,
): string {
  if (!isRecord(value)) return '';
  return [
    borderLine(value.top, x, y, x + width, y),
    borderLine(value.right, x + width, y, x + width, y + height),
    borderLine(value.bottom, x, y + height, x + width, y + height),
    borderLine(value.left, x, y, x, y + height),
  ].join('');
}

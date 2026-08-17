import type { Fill } from './types';
import { svgColor, svgNumber } from './render-svg-values';

export interface PptxSvgGradientPaint {
  definition: string;
  value: string;
}

function gradientOffset(value: string): string | null {
  const match = /^(-?\d+(?:\.\d+)?)%$/.exec(value);
  if (match === null) return null;
  const offset = Number(match[1]);
  return Number.isFinite(offset) && offset >= 0 && offset <= 100
    ? `${svgNumber(offset)}%`
    : null;
}

export function svgLinearGradientPaint(
  fill: Fill,
  id: string,
): PptxSvgGradientPaint | null {
  if (
    fill.type !== 'gradient' ||
    fill.value.path !== 'line' ||
    !Number.isFinite(fill.value.rot) ||
    fill.value.colors.length < 2 ||
    !/^pptx-gradient-[1-9]\d*-[1-9]\d*$/.test(id)
  ) {
    return null;
  }
  const stops = [];
  let previousOffset = -1;
  for (const stop of fill.value.colors) {
    const offset = gradientOffset(stop.pos);
    const color = svgColor(stop.color);
    const numericOffset =
      offset === null ? Number.NaN : Number(offset.slice(0, -1));
    if (offset === null || color === null || numericOffset < previousOffset) {
      return null;
    }
    previousOffset = numericOffset;
    stops.push(`<stop offset="${offset}" stop-color="${color}"/>`);
  }
  const rotation = svgNumber(fill.value.rot) as string;
  return {
    definition: `<linearGradient id="${id}" x1="0" y1="0" x2="1" y2="0" gradientTransform="rotate(${rotation} .5 .5)">${stops.join('')}</linearGradient>`,
    value: `url(#${id})`,
  };
}

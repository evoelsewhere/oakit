import { escapeSvgText, plainTextFromPowerPointHtml } from './render-text';
import { svgColor, svgNumber } from './render-svg-values';

export function renderPptxSvgTableText(
  value: Record<string, unknown>,
  width: number,
  height: number,
): string {
  const source =
    typeof value.text === 'string'
      ? plainTextFromPowerPointHtml(value.text)
      : '';
  if (source === '') return '';
  const lines = source.split('\n');
  const lineHeight = 13;
  const contentHeight = lines.length * lineHeight;
  const start =
    value.vAlign === 'mid'
      ? Math.max(0, (height - contentHeight) / 2)
      : value.vAlign === 'down'
        ? Math.max(0, height - contentHeight - 2)
        : 2;
  const color = svgColor(value.fontColor) ?? '#111827';
  const weight = value.fontBold === true ? ' font-weight="700"' : '';
  const text = lines
    .map(
      (line, index) =>
        `<text x="4" y="${svgNumber(start + 11 + index * lineHeight)}" font-family="sans-serif" font-size="11" fill="${color}"${weight}>${escapeSvgText(line)}</text>`,
    )
    .join('');
  return `<svg x="0" y="0" width="${svgNumber(width)}" height="${svgNumber(height)}" overflow="hidden">${text}</svg>`;
}

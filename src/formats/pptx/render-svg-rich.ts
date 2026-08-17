import { escapeSvgText, plainTextFromPowerPointHtml } from './render-text';
import type { PptxRenderWarningCode } from './render-types';
import { renderPptxSvgTable } from './render-svg-table';
import {
  embeddedRasterDataUri,
  svgBox,
  svgColor,
  svgNumber,
  type PptxSvgBox,
} from './render-svg-values';
import type { Chart, Diagram, Element, Math as MathElement } from './types';

export interface PptxSvgRichElement {
  body: string;
  warningCode: PptxRenderWarningCode;
  warningMessage: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Object(value) === value;
}

function isFiniteNumber(value: unknown): value is number {
  return Number.isFinite(value);
}

function label(text: unknown, fallback: string): string {
  if (typeof text !== 'string') return fallback;
  const plain = plainTextFromPowerPointHtml(text);
  return plain === '' ? fallback : plain;
}

function placeholder(box: PptxSvgBox, text: string): string {
  return `<rect x="0" y="0" width="${svgNumber(box.width)}" height="${svgNumber(box.height)}" fill="#f3f4f6" stroke="#9ca3af" stroke-dasharray="4 3"/><text x="4" y="16" font-family="sans-serif" font-size="12" fill="#374151">${escapeSvgText(text)}</text>`;
}

function chartValues(chart: Chart): number[] {
  if (!Array.isArray(chart.data)) return [];
  const values: number[] = [];
  for (const item of chart.data) {
    if (Array.isArray(item)) {
      for (const value of item) {
        if (isFiniteNumber(value)) values.push(value);
      }
      continue;
    }
    if (!isRecord(item) || !Array.isArray(item.values)) continue;
    for (const point of item.values) {
      if (!isRecord(point)) continue;
      const value = point.y;
      if (isFiniteNumber(value)) values.push(value);
    }
  }
  return values;
}

function chartBody(chart: Chart, box: PptxSvgBox): string {
  const values = chartValues(chart);
  if (values.length === 0) return placeholder(box, 'Chart data unavailable');
  const maximum = values.reduce(
    (current, value) => Math.max(current, Math.abs(value)),
    1,
  );
  const availableHeight = Math.max(box.height - 18, 1);
  const band = box.width / values.length;
  const bars = values
    .map((value, index) => {
      const height = (Math.abs(value) / maximum) * availableHeight;
      const color =
        svgColor(chart.colors[index % chart.colors.length]) ?? '#4f46e5';
      return `<rect x="${svgNumber(index * band + 1)}" y="${svgNumber(box.height - height)}" width="${svgNumber(Math.max(band - 2, 0.5))}" height="${svgNumber(height)}" fill="${color}"/>`;
    })
    .join('');
  return `<rect x="0" y="0" width="${svgNumber(box.width)}" height="${svgNumber(box.height)}" fill="#ffffff" stroke="#d1d5db"/>${bars}<text x="4" y="14" font-family="sans-serif" font-size="11" fill="#374151">${escapeSvgText(chart.chartType)}</text>`;
}

function diagramBody(diagram: Diagram, box: PptxSvgBox): string {
  if (!Array.isArray(diagram.elements) || diagram.elements.length === 0) {
    return placeholder(box, 'Diagram');
  }
  return diagram.elements
    .map((element) => {
      const childBox = svgBox(element);
      if (childBox === null) return '';
      const content = label(
        element.content,
        element.type === 'shape' ? element.name : '',
      );
      return `<g transform="translate(${svgNumber(childBox.left)} ${svgNumber(childBox.top)})"><rect x="0" y="0" width="${svgNumber(childBox.width)}" height="${svgNumber(childBox.height)}" rx="4" fill="#eef2ff" stroke="#818cf8"/><text x="4" y="14" font-family="sans-serif" font-size="11" fill="#312e81">${escapeSvgText(content)}</text></g>`;
    })
    .join('');
}

function mathBody(element: MathElement, box: PptxSvgBox): string {
  const source = embeddedRasterDataUri(element.picBase64);
  if (source !== null) {
    return `<image x="0" y="0" width="${svgNumber(box.width)}" height="${svgNumber(box.height)}" preserveAspectRatio="xMidYMid meet" href="${source}"/>`;
  }
  const text = label(element.text, '');
  const fallback =
    text !== ''
      ? text
      : typeof element.latex === 'string' && element.latex !== ''
        ? element.latex
        : 'Math';
  return placeholder(box, fallback);
}

export function renderPptxSvgRichElement(
  element: Element,
  box: PptxSvgBox,
): PptxSvgRichElement | null {
  switch (element.type) {
    case 'table':
      return {
        body: renderPptxSvgTable(element, box),
        warningCode: 'approximate-table',
        warningMessage:
          'The preview preserves table text and cells with simplified sizing and styling.',
      };
    case 'chart':
      return {
        body: chartBody(element, box),
        warningCode: 'approximate-chart',
        warningMessage:
          'The preview visualizes chart values with simplified portable bars.',
      };
    case 'diagram':
      return {
        body: diagramBody(element, box),
        warningCode: 'approximate-diagram',
        warningMessage:
          'The preview preserves diagram labels with simplified nodes and styling.',
      };
    case 'math':
      return {
        body: mathBody(element, box),
        warningCode: 'approximate-math',
        warningMessage:
          'The preview uses an embedded math image or a portable text fallback.',
      };
    case 'audio':
      return {
        body: placeholder(box, 'Audio'),
        warningCode: 'approximate-media',
        warningMessage:
          'The preview represents audio as a labeled non-interactive placeholder.',
      };
    case 'video':
      return {
        body: placeholder(box, 'Video'),
        warningCode: 'approximate-media',
        warningMessage:
          'The preview represents video as a labeled non-interactive placeholder.',
      };
    default:
      return null;
  }
}

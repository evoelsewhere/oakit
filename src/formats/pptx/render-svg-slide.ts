import type { Element, Fill, Image, PptxSlide, Shape, Text } from './types';
import { escapeSvgText, plainTextFromPowerPointHtml } from './render-text';
import type { PptxRenderWarning } from './render-types';
import { renderPptxSvgRichElement } from './render-svg-rich';
import {
  embeddedRasterDataUri,
  svgBox,
  svgColor,
  svgDashArray,
  svgNumber,
  type PptxSvgBox,
} from './render-svg-values';

export interface PptxSvgSlideSource {
  source: string;
  warnings: readonly PptxRenderWarning[];
}

export interface PptxSvgSlideSourceOptions {
  outputHeight: number;
  outputWidth: number;
  slideNumber: number;
  sourceHeight: number;
  sourceWidth: number;
}

interface RenderContext {
  slideNumber: number;
  warnings: PptxRenderWarning[];
}

function warning(
  context: RenderContext,
  code: PptxRenderWarning['code'],
  message: string,
  element: Element,
): void {
  context.warnings.push({
    code,
    ...(typeof element.id === 'string' && element.id !== ''
      ? { elementId: element.id }
      : {}),
    message,
    slideNumber: context.slideNumber,
  });
}

function paint(
  fill: Fill | null,
  fallback: string,
  context: RenderContext,
  element?: Element,
): string {
  if (fill === null) return 'none';
  if (fill.type === 'color') {
    const color = svgColor(fill.value);
    if (color !== null) return color;
  }
  if (element !== undefined) {
    warning(
      context,
      'approximate-fill',
      'The preview substituted a fill that SVG cannot safely reproduce yet.',
      element,
    );
  }
  return fallback;
}

function strokeAttributes(element: Shape | Text | Image): string {
  const width =
    Number.isFinite(element.borderWidth) && element.borderWidth > 0
      ? svgNumber(element.borderWidth)
      : null;
  const color = svgColor(element.borderColor);
  if (width === null || color === null) return 'stroke="none"';
  const dash = svgDashArray(element.borderStrokeDasharray);
  return `stroke="${color}" stroke-width="${width}"${
    dash === null ? '' : ` stroke-dasharray="${dash}"`
  }`;
}

function localTransform(
  box: PptxSvgBox,
  rotate: unknown,
  flipH: unknown,
  flipV: unknown,
): string {
  const transforms = [
    `translate(${svgNumber(box.left)} ${svgNumber(box.top)})`,
  ];
  const centerX = svgNumber(box.width / 2);
  const centerY = svgNumber(box.height / 2);
  if (Number.isFinite(rotate) && rotate !== 0) {
    transforms.push(`rotate(${svgNumber(rotate)} ${centerX} ${centerY})`);
  }
  if (flipH === true || flipV === true) {
    transforms.push(
      `translate(${flipH === true ? svgNumber(box.width) : '0'} ${
        flipV === true ? svgNumber(box.height) : '0'
      }) scale(${flipH === true ? '-1' : '1'} ${flipV === true ? '-1' : '1'})`,
    );
  }
  return transforms.join(' ');
}

function textLines(content: unknown): string[] {
  if (typeof content !== 'string') return [];
  const text = plainTextFromPowerPointHtml(content);
  return text === '' ? [] : text.split('\n');
}

function textBody(content: unknown, box: PptxSvgBox): string {
  const lines = textLines(content);
  if (lines.length === 0) return '';
  const width = svgNumber(box.width);
  const height = svgNumber(box.height);
  const spans = lines
    .map(
      (line, index) =>
        `<tspan x="4" dy="${index === 0 ? '14' : '16'}">${escapeSvgText(
          line,
        )}</tspan>`,
    )
    .join('');
  return `<svg x="0" y="0" width="${width}" height="${height}" overflow="hidden"><text font-family="sans-serif" font-size="12" fill="#111827">${spans}</text></svg>`;
}

function boxPlaceholder(box: PptxSvgBox, label: string): string {
  const width = svgNumber(box.width);
  const height = svgNumber(box.height);
  return `<rect x="0" y="0" width="${width}" height="${height}" fill="#f3f4f6" stroke="#9ca3af" stroke-dasharray="4 3"/><text x="4" y="16" font-family="sans-serif" font-size="12" fill="#374151">${escapeSvgText(label)}</text>`;
}

function renderText(
  element: Text,
  box: PptxSvgBox,
  context: RenderContext,
): string {
  warning(
    context,
    'font-substitution',
    'The preview uses a portable sans-serif font instead of the authored font.',
    element,
  );
  const width = svgNumber(box.width);
  const height = svgNumber(box.height);
  const fill = paint(element.fill, 'none', context, element);
  return `<g transform="${localTransform(box, element.rotate, element.isFlipH, element.isFlipV)}"><rect x="0" y="0" width="${width}" height="${height}" fill="${fill}" ${strokeAttributes(element)}/>${textBody(element.content, box)}</g>`;
}

function shapeGeometry(
  element: Shape,
  box: PptxSvgBox,
  context: RenderContext,
): string {
  const width = svgNumber(box.width);
  const height = svgNumber(box.height);
  const fill =
    element.strokeOnly === true
      ? 'none'
      : paint(element.fill, 'none', context, element);
  const style = `fill="${fill}" ${strokeAttributes(element)}`;
  if (element.shapType === 'ellipse') {
    return `<ellipse cx="${svgNumber(box.width / 2)}" cy="${svgNumber(box.height / 2)}" rx="${svgNumber(box.width / 2)}" ry="${svgNumber(box.height / 2)}" ${style}/>`;
  }
  if (element.shapType === 'line') {
    return `<line x1="0" y1="0" x2="${width}" y2="${height}" ${style}/>`;
  }
  if (element.shapType === 'lineInv') {
    return `<line x1="0" y1="${height}" x2="${width}" y2="0" ${style}/>`;
  }
  const pathViewBox = element.pathViewBox;
  const viewBox =
    pathViewBox === undefined
      ? null
      : svgBox({
          height: pathViewBox.height,
          left: pathViewBox.x,
          top: pathViewBox.y,
          width: pathViewBox.width,
        });
  if (typeof element.path === 'string' && element.path !== '' && viewBox) {
    return `<svg x="0" y="0" width="${width}" height="${height}" viewBox="${svgNumber(viewBox.left)} ${svgNumber(viewBox.top)} ${svgNumber(viewBox.width)} ${svgNumber(viewBox.height)}" preserveAspectRatio="none"><path d="${escapeSvgText(element.path)}" ${style}/></svg>`;
  }
  if (element.shapType !== 'rect' && element.shapType !== 'roundRect') {
    warning(
      context,
      'approximate-shape',
      `The preview represents PowerPoint shape ${String(element.shapType)} as a rectangle.`,
      element,
    );
  }
  const radius = element.shapType === 'roundRect' ? '6' : '0';
  return `<rect x="0" y="0" width="${width}" height="${height}" rx="${radius}" ${style}/>`;
}

function renderShape(
  element: Shape,
  box: PptxSvgBox,
  context: RenderContext,
): string {
  const body = textBody(element.content, box);
  if (body !== '') {
    warning(
      context,
      'font-substitution',
      'The preview uses a portable sans-serif font instead of the authored font.',
      element,
    );
  }
  return `<g transform="${localTransform(box, element.rotate, element.isFlipH, element.isFlipV)}">${shapeGeometry(element, box, context)}${body}</g>`;
}

function renderImage(
  element: Image,
  box: PptxSvgBox,
  context: RenderContext,
): string {
  const transform = localTransform(
    box,
    element.rotate,
    element.isFlipH,
    element.isFlipV,
  );
  const width = svgNumber(box.width);
  const height = svgNumber(box.height);
  const source = embeddedRasterDataUri(element.base64);
  if (source === null) {
    warning(
      context,
      'missing-media',
      'The preview omitted an image because no safe embedded raster source was available.',
      element,
    );
    return `<g transform="${transform}">${boxPlaceholder(box, 'Image unavailable')}</g>`;
  }
  if (element.rect !== undefined || element.filters !== undefined) {
    warning(
      context,
      'approximate-media',
      'The preview preserves image bytes but may approximate crop or filter effects.',
      element,
    );
  }
  return `<g transform="${transform}"><image x="0" y="0" width="${width}" height="${height}" preserveAspectRatio="xMidYMid meet" href="${source}"/><rect x="0" y="0" width="${width}" height="${height}" fill="none" ${strokeAttributes(element)}/></g>`;
}

function renderElement(element: Element, context: RenderContext): string {
  const box = svgBox(element);
  if (box === null) {
    warning(
      context,
      'approximate-shape',
      'The preview skipped an element with invalid geometry.',
      element,
    );
    return '';
  }
  switch (element.type) {
    case 'text':
      return renderText(element, box, context);
    case 'shape':
      return renderShape(element, box, context);
    case 'image':
      return renderImage(element, box, context);
    case 'group': {
      const children = element.elements
        .map((child) => renderElement(child, context))
        .join('');
      return `<g transform="${localTransform(box, element.rotate, element.isFlipH, element.isFlipV)}">${children}</g>`;
    }
    default: {
      const rich = renderPptxSvgRichElement(element, box);
      if (rich !== null) {
        warning(context, rich.warningCode, rich.warningMessage, element);
        return `<g transform="${localTransform(box, 0, false, false)}">${rich.body}</g>`;
      }
      warning(
        context,
        'approximate-shape',
        `The preview represents PowerPoint element ${element.type} as a placeholder.`,
        element,
      );
      return `<g transform="${localTransform(box, 0, false, false)}">${boxPlaceholder(box, element.type)}</g>`;
    }
  }
}

export function renderPptxSvgSlideSource(
  slide: PptxSlide,
  options: PptxSvgSlideSourceOptions,
): PptxSvgSlideSource {
  const context: RenderContext = {
    slideNumber: options.slideNumber,
    warnings: [],
  };
  const background = paint(slide.fill, '#ffffff', context);
  const elements = [...slide.layoutElements, ...slide.elements]
    .map((element) => renderElement(element, context))
    .join('');
  const sourceWidth = svgNumber(options.sourceWidth);
  const sourceHeight = svgNumber(options.sourceHeight);
  const outputWidth = svgNumber(options.outputWidth);
  const outputHeight = svgNumber(options.outputHeight);
  const source = `<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg" role="img" width="${outputWidth}" height="${outputHeight}" viewBox="0 0 ${sourceWidth} ${sourceHeight}"><title>PowerPoint slide ${options.slideNumber}</title><rect width="${sourceWidth}" height="${sourceHeight}" fill="${background}"/>${elements}</svg>`;
  return { source, warnings: context.warnings };
}

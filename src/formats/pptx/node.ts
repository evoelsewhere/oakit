import { Buffer } from 'node:buffer';

import { Resvg } from '@resvg/resvg-js';

import { PptxRenderError } from './render-error';
import { renderPptxDocumentToSvg, renderPptxToSvg } from './render-svg';
import type {
  PptxInputRenderOptions,
  PptxPngRenderResult,
  PptxRenderedPngSlide,
  PptxRenderedSvgSlide,
  PptxSvgRenderResult,
} from './render-types';
import type { PptxDocument, PptxInput } from './types';

function rasterizeSlide(slide: PptxRenderedSvgSlide): PptxRenderedPngSlide {
  try {
    const data = Uint8Array.from(
      new Resvg(Buffer.from(slide.data)).render().asPng(),
    );
    return {
      data,
      format: 'png',
      height: slide.height,
      mimeType: 'image/png',
      slideNumber: slide.slideNumber,
      warnings: slide.warnings,
      width: slide.width,
    };
  } catch (cause) {
    throw new PptxRenderError(
      'rasterization-failed',
      `PowerPoint slide ${slide.slideNumber} could not be rasterized as PNG`,
      { cause },
    );
  }
}

/** Rasterize self-contained SVG slide bytes into PNG without an Office runtime. */
export function rasterizePptxSvgResult(
  result: PptxSvgRenderResult,
): PptxPngRenderResult {
  return { slides: result.slides.map(rasterizeSlide) };
}

/** Render a parsed PowerPoint document into PNG bytes in Node.js. */
export function renderPptxDocumentToPng(
  document: PptxDocument,
  options: PptxInputRenderOptions = {},
): PptxPngRenderResult {
  return rasterizePptxSvgResult(renderPptxDocumentToSvg(document, options));
}

/** Open and render a PowerPoint package into PNG bytes in Node.js. */
export async function renderPptxToPng(
  input: PptxInput,
  options: PptxInputRenderOptions = {},
): Promise<PptxPngRenderResult> {
  return rasterizePptxSvgResult(await renderPptxToSvg(input, options));
}

export { PptxRenderError } from './render-error';
export type * from './render-types';
export type { PptxDocument, PptxInput } from './types';

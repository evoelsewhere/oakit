import { PptxRenderError } from './render-error';
import { resolvePptxRenderRequest } from './render-request';
import { renderPptxSvgSlideSource } from './render-svg-slide';
import type {
  PptxRenderedSvgSlide,
  PptxRenderOptions,
  PptxSvgRenderResult,
} from './render-types';
import type { PptxDocument } from './types';

const UTF8_ENCODER = new TextEncoder();

/** Render a parsed PowerPoint document into portable, self-contained SVG bytes. */
export function renderPptxDocumentToSvg(
  document: PptxDocument,
  options: PptxRenderOptions = {},
): PptxSvgRenderResult {
  const request = resolvePptxRenderRequest(document, options);
  const slides: PptxRenderedSvgSlide[] = request.slides.map(
    ({ slide, slideNumber }) => {
      const rendered = renderPptxSvgSlideSource(slide, {
        outputHeight: request.height,
        outputWidth: request.width,
        slideNumber,
        sourceHeight: document.size.height,
        sourceWidth: document.size.width,
      });
      const data = UTF8_ENCODER.encode(rendered.source);
      if (data.byteLength > request.limits.maxSvgBytes) {
        throw new PptxRenderError(
          'resource-limit-exceeded',
          `PowerPoint slide ${slideNumber} SVG exceeds the ${request.limits.maxSvgBytes} byte limit`,
        );
      }
      return {
        data,
        format: 'svg',
        height: request.height,
        mimeType: 'image/svg+xml',
        slideNumber,
        warnings: rendered.warnings,
        width: request.width,
      };
    },
  );
  return { slides };
}

export type PptxRenderWarningCode =
  | 'approximate-chart'
  | 'approximate-diagram'
  | 'approximate-fill'
  | 'approximate-math'
  | 'approximate-media'
  | 'approximate-shape'
  | 'approximate-table'
  | 'font-substitution'
  | 'missing-media';

export interface PptxRenderWarning {
  code: PptxRenderWarningCode;
  elementId?: string;
  message: string;
  slideNumber: number;
}

export interface PptxRenderLimits {
  /** Maximum number of slides rendered by one request. */
  maxSlides?: number;
  /** Maximum number of elements traversed on one slide, including groups. */
  maxElementsPerSlide?: number;
  /** Maximum scale multiplier applied to source slide dimensions. */
  maxScale?: number;
  /** Maximum output pixels for one slide. */
  maxOutputPixels?: number;
  /** Maximum encoded PNG bytes for one slide in the Node.js renderer. */
  maxPngBytes?: number;
  /** Maximum encoded SVG bytes for one slide. */
  maxSvgBytes?: number;
}

export interface PptxRenderOptions {
  limits?: PptxRenderLimits;
  /** Output pixels per source point. Defaults to 1. */
  scale?: number;
  /** One-based slide numbers. Omit to render every slide in source order. */
  slideNumbers?: readonly number[];
}

export interface PptxInputRenderOptions extends PptxRenderOptions {
  /** Resource limits applied while opening the PowerPoint package. */
  parseLimits?: import('./types').PptxResourceLimits;
}

export interface PptxRenderedSvgSlide {
  data: Uint8Array;
  format: 'svg';
  height: number;
  mimeType: 'image/svg+xml';
  slideNumber: number;
  warnings: readonly PptxRenderWarning[];
  width: number;
}

export interface PptxRenderedPngSlide {
  data: Uint8Array;
  format: 'png';
  height: number;
  mimeType: 'image/png';
  slideNumber: number;
  warnings: readonly PptxRenderWarning[];
  width: number;
}

export interface PptxSvgRenderResult {
  slides: readonly PptxRenderedSvgSlide[];
}

export interface PptxPngRenderResult {
  slides: readonly PptxRenderedPngSlide[];
}

export type PptxRenderErrorCode =
  | 'invalid-document'
  | 'invalid-option'
  | 'rasterization-failed'
  | 'resource-limit-exceeded'
  | 'slide-not-found';

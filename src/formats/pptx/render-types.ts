export type PptxRenderWarningCode =
  | 'approximate-chart'
  | 'approximate-fill'
  | 'approximate-media'
  | 'approximate-shape'
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

export interface PptxRenderedSvgSlide {
  data: Uint8Array;
  format: 'svg';
  height: number;
  mimeType: 'image/svg+xml';
  slideNumber: number;
  warnings: readonly PptxRenderWarning[];
  width: number;
}

export interface PptxSvgRenderResult {
  slides: readonly PptxRenderedSvgSlide[];
}

export type PptxRenderErrorCode =
  | 'invalid-document'
  | 'invalid-option'
  | 'resource-limit-exceeded'
  | 'slide-not-found';

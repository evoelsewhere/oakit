import { PptxRenderError } from './render-error';
import type { PptxRenderLimits } from './render-types';

const DEFAULT_LIMITS: Required<PptxRenderLimits> = Object.freeze({
  maxElementsPerSlide: 10_000,
  maxOutputPixels: 32 * 1024 * 1024,
  maxPngBytes: 256 * 1024 * 1024,
  maxScale: 8,
  maxSlides: 1_000,
  maxSvgBytes: 128 * 1024 * 1024,
});

export function defaultPptxRenderLimits(): Required<PptxRenderLimits> {
  return { ...DEFAULT_LIMITS };
}

function positiveSafeInteger(name: string, value: number): number {
  if (!Number.isSafeInteger(value) || value <= 0) {
    throw new PptxRenderError(
      'invalid-option',
      `PowerPoint render limit ${name} must be a positive safe integer`,
    );
  }
  return value;
}

function positiveFiniteNumber(name: string, value: number): number {
  if (!Number.isFinite(value) || value <= 0) {
    throw new PptxRenderError(
      'invalid-option',
      `PowerPoint render limit ${name} must be a positive finite number`,
    );
  }
  return value;
}

export function resolvePptxRenderLimits(
  limits: PptxRenderLimits = {},
): Required<PptxRenderLimits> {
  const resolved = { ...DEFAULT_LIMITS, ...limits };
  return {
    maxElementsPerSlide: positiveSafeInteger(
      'maxElementsPerSlide',
      resolved.maxElementsPerSlide,
    ),
    maxOutputPixels: positiveSafeInteger(
      'maxOutputPixels',
      resolved.maxOutputPixels,
    ),
    maxPngBytes: positiveSafeInteger('maxPngBytes', resolved.maxPngBytes),
    maxScale: positiveFiniteNumber('maxScale', resolved.maxScale),
    maxSlides: positiveSafeInteger('maxSlides', resolved.maxSlides),
    maxSvgBytes: positiveSafeInteger('maxSvgBytes', resolved.maxSvgBytes),
  };
}

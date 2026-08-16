import { PptxRenderError } from './render-error';
import { resolvePptxRenderLimits } from './render-limits';
import type { PptxRenderOptions } from './render-types';
import type { Element, PptxDocument, PptxSlide } from './types';

export interface ResolvedPptxRenderSlide {
  slide: PptxSlide;
  slideNumber: number;
}

export interface ResolvedPptxRenderRequest {
  height: number;
  limits: ReturnType<typeof resolvePptxRenderLimits>;
  scale: number;
  slides: readonly ResolvedPptxRenderSlide[];
  width: number;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function positiveFinite(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value) && value > 0;
}

function isNumberArray(value: unknown): value is readonly number[] {
  return Array.isArray(value);
}

function assertDocument(value: unknown): asserts value is PptxDocument {
  if (!isRecord(value) || !isRecord(value.size)) {
    throw new PptxRenderError(
      'invalid-document',
      'PowerPoint render input must contain a size object',
    );
  }
  if (!positiveFinite(value.size.width) || !positiveFinite(value.size.height)) {
    throw new PptxRenderError(
      'invalid-document',
      'PowerPoint render input must have positive finite slide dimensions',
    );
  }
  if (!Array.isArray(value.slides)) {
    throw new PptxRenderError(
      'invalid-document',
      'PowerPoint render input must contain a slides array',
    );
  }
}

function requestedSlideNumbers(
  document: PptxDocument,
  requested: readonly number[] | undefined,
): number[] {
  if (requested === undefined) {
    return document.slides.map((_slide, index) => index + 1);
  }
  if (!isNumberArray(requested)) {
    throw new PptxRenderError(
      'invalid-option',
      'PowerPoint render slideNumbers must be an array',
    );
  }
  return [...requested];
}

function selectSlides(
  document: PptxDocument,
  requested: readonly number[] | undefined,
  maxSlides: number,
): ResolvedPptxRenderSlide[] {
  const numbers = requestedSlideNumbers(document, requested);
  if (numbers.length > maxSlides) {
    throw new PptxRenderError(
      'resource-limit-exceeded',
      `PowerPoint render request exceeds the ${maxSlides} slide limit`,
    );
  }

  const seen = new Set<number>();
  return numbers.map((slideNumber) => {
    if (!Number.isSafeInteger(slideNumber) || slideNumber <= 0) {
      throw new PptxRenderError(
        'invalid-option',
        'PowerPoint render slide numbers must be positive safe integers',
      );
    }
    if (seen.has(slideNumber)) {
      throw new PptxRenderError(
        'invalid-option',
        `PowerPoint render slide number ${slideNumber} is duplicated`,
      );
    }
    seen.add(slideNumber);
    const slide = document.slides[slideNumber - 1];
    if (slide === undefined) {
      throw new PptxRenderError(
        'slide-not-found',
        `PowerPoint slide ${slideNumber} does not exist`,
      );
    }
    return { slide, slideNumber };
  });
}

function childElements(element: Element): Element[] {
  if (element.type !== 'group') return [];
  if (!Array.isArray(element.elements)) {
    throw new PptxRenderError(
      'invalid-document',
      'PowerPoint group elements must be an array',
    );
  }
  return element.elements;
}

function assertSlideElements(slide: PptxSlide, maximum: number): void {
  if (!Array.isArray(slide.elements) || !Array.isArray(slide.layoutElements)) {
    throw new PptxRenderError(
      'invalid-document',
      'PowerPoint slide elements and layoutElements must be arrays',
    );
  }
  const queue: unknown[] = [...slide.layoutElements, ...slide.elements];
  const seen = new WeakSet<object>();
  let count = 0;
  for (const element of queue) {
    if (!isRecord(element) || typeof element.type !== 'string') {
      throw new PptxRenderError(
        'invalid-document',
        'PowerPoint render elements must be typed objects',
      );
    }
    if (seen.has(element)) {
      throw new PptxRenderError(
        'invalid-document',
        'PowerPoint render elements must not contain cycles or shared objects',
      );
    }
    seen.add(element);
    count += 1;
    if (count > maximum) {
      throw new PptxRenderError(
        'resource-limit-exceeded',
        `PowerPoint slide exceeds the ${maximum} element limit`,
      );
    }
    queue.push(...childElements(element as unknown as Element));
  }
}

function renderScale(value: number | undefined, maximum: number): number {
  const scale = value ?? 1;
  if (!positiveFinite(scale)) {
    throw new PptxRenderError(
      'invalid-option',
      'PowerPoint render scale must be a positive finite number',
    );
  }
  if (scale > maximum) {
    throw new PptxRenderError(
      'resource-limit-exceeded',
      `PowerPoint render scale exceeds the ${maximum} limit`,
    );
  }
  return scale;
}

function outputDimension(source: number, scale: number): number {
  const output = Math.round(source * scale);
  if (!Number.isSafeInteger(output) || output <= 0) {
    throw new PptxRenderError(
      'resource-limit-exceeded',
      'PowerPoint render dimensions exceed the safe output range',
    );
  }
  return output;
}

export function resolvePptxRenderRequest(
  value: PptxDocument,
  options: PptxRenderOptions = {},
): ResolvedPptxRenderRequest {
  assertDocument(value);
  const limits = resolvePptxRenderLimits(options.limits);
  const scale = renderScale(options.scale, limits.maxScale);
  const width = outputDimension(value.size.width, scale);
  const height = outputDimension(value.size.height, scale);
  if (width > Math.floor(limits.maxOutputPixels / height)) {
    throw new PptxRenderError(
      'resource-limit-exceeded',
      `PowerPoint render output exceeds the ${limits.maxOutputPixels} pixel limit`,
    );
  }
  const slides = selectSlides(value, options.slideNumbers, limits.maxSlides);
  slides.forEach(({ slide }) =>
    assertSlideElements(slide, limits.maxElementsPerSlide),
  );
  return { height, limits, scale, slides, width };
}

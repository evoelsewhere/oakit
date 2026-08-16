import { describe, expect, it } from 'vitest';

import { PptxRenderError } from '../../src/formats/pptx/render-error';
import {
  defaultPptxRenderLimits,
  resolvePptxRenderLimits,
} from '../../src/formats/pptx/render-limits';

describe('PowerPoint render limits', () => {
  it('returns an isolated copy of every default', () => {
    const first = defaultPptxRenderLimits();
    const second = defaultPptxRenderLimits();

    expect(first).toEqual({
      maxElementsPerSlide: 10_000,
      maxOutputPixels: 33_554_432,
      maxPngBytes: 268_435_456,
      maxScale: 8,
      maxSlides: 1_000,
      maxSvgBytes: 134_217_728,
    });
    expect(first).not.toBe(second);
    first.maxSlides = 1;
    expect(defaultPptxRenderLimits().maxSlides).toBe(1_000);
  });

  it('merges caller overrides without changing unspecified defaults', () => {
    expect(
      resolvePptxRenderLimits({
        maxElementsPerSlide: 7,
        maxOutputPixels: 8,
        maxPngBytes: 10,
        maxScale: 1.5,
        maxSlides: 2,
        maxSvgBytes: 9,
      }),
    ).toEqual({
      maxElementsPerSlide: 7,
      maxOutputPixels: 8,
      maxPngBytes: 10,
      maxScale: 1.5,
      maxSlides: 2,
      maxSvgBytes: 9,
    });
    expect(resolvePptxRenderLimits({ maxSlides: 3 })).toEqual({
      ...defaultPptxRenderLimits(),
      maxSlides: 3,
    });
  });

  it.each([
    ['maxElementsPerSlide', 0],
    ['maxElementsPerSlide', 1.5],
    ['maxOutputPixels', -1],
    ['maxOutputPixels', Number.MAX_SAFE_INTEGER + 1],
    ['maxPngBytes', 1.5],
    ['maxSlides', Number.NaN],
    ['maxSvgBytes', Number.POSITIVE_INFINITY],
  ] as const)('rejects invalid integer limit %s=%s', (name, value) => {
    expect(() => resolvePptxRenderLimits({ [name]: value })).toThrow(
      new PptxRenderError(
        'invalid-option',
        `PowerPoint render limit ${name} must be a positive safe integer`,
      ),
    );
  });

  it.each([0, -1, Number.NaN, Number.POSITIVE_INFINITY])(
    'rejects invalid maximum scale %s',
    (maxScale) => {
      expect(() => resolvePptxRenderLimits({ maxScale })).toThrow(
        new PptxRenderError(
          'invalid-option',
          'PowerPoint render limit maxScale must be a positive finite number',
        ),
      );
    },
  );

  it('exposes stable typed error metadata and an optional cause', () => {
    const cause = new Error('root');
    const error = new PptxRenderError('invalid-document', 'broken', { cause });

    expect(error).toMatchObject({
      cause,
      code: 'invalid-document',
      message: 'broken',
      name: 'PptxRenderError',
    });
    expect(error).toBeInstanceOf(Error);
  });
});

import { describe, expect, it } from 'vitest';

import { PptxRenderError } from '../../src/formats/pptx/render-error';
import { resolvePptxRenderRequest } from '../../src/formats/pptx/render-request';
import type {
  Element,
  PptxDocument,
  PptxSlide,
} from '../../src/formats/pptx/types';

function slide(elements: Element[] = []): PptxSlide {
  return {
    elements,
    fill: { type: 'color', value: '#ffffff' },
    layoutElements: [],
    note: '',
  };
}

function document(slides: PptxSlide[] = [slide(), slide()]): PptxDocument {
  return {
    size: { height: 50, width: 100 },
    slides,
    themeColors: [],
    usedFonts: [],
  };
}

function text(id: string): Element {
  return {
    borderColor: '',
    borderStrokeDasharray: '',
    borderType: 'solid',
    borderWidth: 0,
    content: id,
    fill: null,
    height: 10,
    id,
    isFlipH: false,
    isFlipV: false,
    isVertical: false,
    left: 0,
    name: id,
    order: 0,
    rotate: 0,
    top: 0,
    type: 'text',
    vAlign: 'top',
    width: 10,
    wrap: true,
  };
}

function group(id: string, elements: Element[]): Element {
  return {
    elements,
    height: 10,
    id,
    isFlipH: false,
    isFlipV: false,
    left: 0,
    order: 0,
    rotate: 0,
    top: 0,
    type: 'group',
    width: 10,
  };
}

describe('PowerPoint render requests', () => {
  it('selects every slide by default and resolves output pixels', () => {
    const result = resolvePptxRenderRequest(document());

    expect(result).toMatchObject({ height: 50, scale: 1, width: 100 });
    expect(result.slides.map(({ slideNumber }) => slideNumber)).toEqual([1, 2]);
  });

  it('preserves an explicit slide order and accepts the exact scale limit', () => {
    const result = resolvePptxRenderRequest(document(), {
      scale: 2,
      slideNumbers: [2, 1],
      limits: { maxScale: 2 },
    });

    expect(result).toMatchObject({ height: 100, scale: 2, width: 200 });
    expect(result.slides.map(({ slideNumber }) => slideNumber)).toEqual([2, 1]);
  });

  it('allows an explicit empty selection', () => {
    expect(
      resolvePptxRenderRequest(document(), { slideNumbers: [] }).slides,
    ).toEqual([]);
  });

  it.each([
    [null, 'PowerPoint render input must contain a size object'],
    [7, 'PowerPoint render input must contain a size object'],
    [
      { slides: [], size: null },
      'PowerPoint render input must contain a size object',
    ],
    [
      { slides: [], size: { height: 1, width: 0 } },
      'PowerPoint render input must have positive finite slide dimensions',
    ],
    [
      { slides: [], size: { height: Number.NaN, width: 1 } },
      'PowerPoint render input must have positive finite slide dimensions',
    ],
    [
      { slides: null, size: { height: 1, width: 1 } },
      'PowerPoint render input must contain a slides array',
    ],
  ])('rejects malformed document %#', (value, message) => {
    expect(() =>
      resolvePptxRenderRequest(value as unknown as PptxDocument),
    ).toThrow(new PptxRenderError('invalid-document', message));
  });

  it.each([0, -1, Number.NaN, Number.POSITIVE_INFINITY])(
    'rejects invalid scale %s',
    (scale) => {
      expect(() => resolvePptxRenderRequest(document(), { scale })).toThrow(
        new PptxRenderError(
          'invalid-option',
          'PowerPoint render scale must be a positive finite number',
        ),
      );
    },
  );

  it('rejects a scale above the configured maximum', () => {
    expect(() =>
      resolvePptxRenderRequest(document(), {
        limits: { maxScale: 1 },
        scale: 1.1,
      }),
    ).toThrow(
      new PptxRenderError(
        'resource-limit-exceeded',
        'PowerPoint render scale exceeds the 1 limit',
      ),
    );
  });

  it('rejects unsafe rounded output dimensions', () => {
    const unsafe = document();
    unsafe.size.width = Number.MAX_VALUE;
    expect(() => resolvePptxRenderRequest(unsafe)).toThrow(
      new PptxRenderError(
        'resource-limit-exceeded',
        'PowerPoint render dimensions exceed the safe output range',
      ),
    );

    const roundedToZero = document();
    roundedToZero.size = { height: 0.1, width: 0.1 };
    expect(() =>
      resolvePptxRenderRequest(roundedToZero, { scale: 0.1 }),
    ).toThrow(
      new PptxRenderError(
        'resource-limit-exceeded',
        'PowerPoint render dimensions exceed the safe output range',
      ),
    );
  });

  it('accepts the exact pixel budget and rejects one pixel below', () => {
    const input = document([slide()]);
    input.size = { height: 2, width: 4 };
    expect(
      resolvePptxRenderRequest(input, { limits: { maxOutputPixels: 8 } }),
    ).toMatchObject({ height: 2, width: 4 });
    expect(() =>
      resolvePptxRenderRequest(input, { limits: { maxOutputPixels: 7 } }),
    ).toThrow(
      new PptxRenderError(
        'resource-limit-exceeded',
        'PowerPoint render output exceeds the 7 pixel limit',
      ),
    );
  });

  it.each([0, -1, 1.5, Number.MAX_SAFE_INTEGER + 1])(
    'rejects invalid slide number %s',
    (slideNumber) => {
      expect(() =>
        resolvePptxRenderRequest(document(), { slideNumbers: [slideNumber] }),
      ).toThrow(
        new PptxRenderError(
          'invalid-option',
          'PowerPoint render slide numbers must be positive safe integers',
        ),
      );
    },
  );

  it('rejects a non-array slide selection at runtime', () => {
    expect(() =>
      resolvePptxRenderRequest(document(), {
        slideNumbers: new Set([1]) as unknown as readonly number[],
      }),
    ).toThrow(
      new PptxRenderError(
        'invalid-option',
        'PowerPoint render slideNumbers must be an array',
      ),
    );
  });

  it('rejects duplicate and absent slide numbers independently', () => {
    expect(() =>
      resolvePptxRenderRequest(document(), { slideNumbers: [1, 1] }),
    ).toThrow(
      new PptxRenderError(
        'invalid-option',
        'PowerPoint render slide number 1 is duplicated',
      ),
    );
    expect(() =>
      resolvePptxRenderRequest(document(), { slideNumbers: [3] }),
    ).toThrow(
      new PptxRenderError(
        'slide-not-found',
        'PowerPoint slide 3 does not exist',
      ),
    );
  });

  it('accepts the exact slide limit and rejects one above it', () => {
    expect(
      resolvePptxRenderRequest(document(), { limits: { maxSlides: 2 } }).slides,
    ).toHaveLength(2);
    expect(() =>
      resolvePptxRenderRequest(document(), { limits: { maxSlides: 1 } }),
    ).toThrow(
      new PptxRenderError(
        'resource-limit-exceeded',
        'PowerPoint render request exceeds the 1 slide limit',
      ),
    );
  });

  it('counts layout, slide, and nested group elements at the exact boundary', () => {
    const inputSlide = slide([group('g', [text('nested')])]);
    inputSlide.layoutElements.push(text('layout'));
    expect(
      resolvePptxRenderRequest(document([inputSlide]), {
        limits: { maxElementsPerSlide: 3 },
      }).slides,
    ).toHaveLength(1);
    expect(() =>
      resolvePptxRenderRequest(document([inputSlide]), {
        limits: { maxElementsPerSlide: 2 },
      }),
    ).toThrow(
      new PptxRenderError(
        'resource-limit-exceeded',
        'PowerPoint slide exceeds the 2 element limit',
      ),
    );
  });

  it.each([
    [{}, 'PowerPoint render elements must be typed objects'],
    [
      { type: 'group', elements: null },
      'PowerPoint group elements must be an array',
    ],
  ])('rejects malformed element %#', (element, message) => {
    const inputSlide = slide([element as unknown as Element]);
    expect(() => resolvePptxRenderRequest(document([inputSlide]))).toThrow(
      new PptxRenderError('invalid-document', message),
    );
  });

  it('rejects malformed slide-owned element arrays', () => {
    const inputSlide = slide();
    inputSlide.layoutElements = null as unknown as Element[];
    expect(() => resolvePptxRenderRequest(document([inputSlide]))).toThrow(
      new PptxRenderError(
        'invalid-document',
        'PowerPoint slide elements and layoutElements must be arrays',
      ),
    );
  });

  it('rejects cycles and shared element objects', () => {
    const cyclic = group('cycle', []);
    if (cyclic.type !== 'group') throw new Error('Expected group');
    cyclic.elements.push(cyclic);
    expect(() =>
      resolvePptxRenderRequest(document([slide([cyclic])]), {
        limits: { maxElementsPerSlide: 3 },
      }),
    ).toThrow(
      new PptxRenderError(
        'invalid-document',
        'PowerPoint render elements must not contain cycles or shared objects',
      ),
    );

    const shared = text('shared');
    expect(() =>
      resolvePptxRenderRequest(document([slide([shared, shared])])),
    ).toThrow(
      new PptxRenderError(
        'invalid-document',
        'PowerPoint render elements must not contain cycles or shared objects',
      ),
    );
  });
});

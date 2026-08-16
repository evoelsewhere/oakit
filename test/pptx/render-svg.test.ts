import { describe, expect, it } from 'vitest';

import { PptxRenderError, renderPptxDocumentToSvg } from '../../src/index';
import type {
  PptxDocument,
  PptxSlide,
  Text,
} from '../../src/formats/pptx/types';

const UTF8_DECODER = new TextDecoder();

function text(content: string): Text {
  return {
    borderColor: '',
    borderStrokeDasharray: '',
    borderType: 'solid',
    borderWidth: 0,
    content,
    fill: null,
    height: 20,
    id: content,
    isFlipH: false,
    isFlipV: false,
    isVertical: false,
    left: 1,
    name: content,
    order: 0,
    rotate: 0,
    top: 2,
    type: 'text',
    vAlign: 'top',
    width: 40,
    wrap: true,
  };
}

function slide(content: string): PptxSlide {
  return {
    elements: [text(content)],
    fill: { type: 'color', value: '#ffffff' },
    layoutElements: [],
    note: '',
  };
}

function document(): PptxDocument {
  return {
    size: { height: 50, width: 100 },
    slides: [slide('First'), slide('Second')],
    themeColors: [],
    usedFonts: [],
  };
}

describe('PowerPoint document SVG rendering', () => {
  it('exposes selected slides as self-contained UTF-8 SVG metadata', () => {
    const result = renderPptxDocumentToSvg(document(), {
      scale: 2,
      slideNumbers: [2],
    });

    expect(result.slides).toHaveLength(1);
    const rendered = result.slides[0];
    expect(rendered).toMatchObject({
      format: 'svg',
      height: 100,
      mimeType: 'image/svg+xml',
      slideNumber: 2,
      width: 200,
    });
    expect(rendered?.data).toBeInstanceOf(Uint8Array);
    expect(UTF8_DECODER.decode(rendered?.data)).toContain(
      '<title>PowerPoint slide 2</title>',
    );
    expect(UTF8_DECODER.decode(rendered?.data)).toContain('>Second</tspan>');
    expect(rendered?.warnings.map(({ code }) => code)).toEqual([
      'font-substitution',
    ]);
  });

  it('renders every slide in source order by default with independent bytes', () => {
    const result = renderPptxDocumentToSvg(document());

    expect(result.slides.map(({ slideNumber }) => slideNumber)).toEqual([1, 2]);
    expect(result.slides[0]?.data).not.toBe(result.slides[1]?.data);
    expect(UTF8_DECODER.decode(result.slides[0]?.data)).toContain(
      '>First</tspan>',
    );
    expect(UTF8_DECODER.decode(result.slides[1]?.data)).toContain(
      '>Second</tspan>',
    );
  });

  it('supports an explicit empty selection without allocating slide output', () => {
    expect(renderPptxDocumentToSvg(document(), { slideNumbers: [] })).toEqual({
      slides: [],
    });
  });

  it('accepts the exact SVG byte budget and rejects one byte below it', () => {
    const input = document();
    input.slides = [slide('Budget')];
    const exact = renderPptxDocumentToSvg(input).slides[0]?.data.byteLength;
    if (exact === undefined) throw new Error('Expected one rendered slide');

    expect(
      renderPptxDocumentToSvg(input, { limits: { maxSvgBytes: exact } })
        .slides[0]?.data.byteLength,
    ).toBe(exact);
    expect(() =>
      renderPptxDocumentToSvg(input, {
        limits: { maxSvgBytes: exact - 1 },
      }),
    ).toThrow(
      new PptxRenderError(
        'resource-limit-exceeded',
        `PowerPoint slide 1 SVG exceeds the ${exact - 1} byte limit`,
      ),
    );
  });
});

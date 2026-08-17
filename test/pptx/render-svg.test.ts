import { describe, expect, it } from 'vitest';

import {
  PptxRenderError,
  renderPptxDocumentToSvg,
  renderPptxToSvg,
} from '../../src/index';
import type {
  PptxDocument,
  PptxInput,
  PptxSlide,
  Text,
} from '../../src/formats/pptx/types';
import { createMinimalPptx } from './fixture';

const UTF8_DECODER = new TextDecoder();

const PACKAGE_INPUTS: [string, (bytes: Uint8Array) => PptxInput][] = [
  ['Uint8Array', (bytes) => bytes],
  ['ArrayBuffer', (bytes) => Uint8Array.from(bytes).buffer],
  ['Blob', (bytes) => new Blob([Uint8Array.from(bytes).buffer])],
];

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

  it('renders rich text color, size, emphasis, and alignment deterministically', () => {
    const element = text(
      '<p style="text-align: left"><span style="color: #F8FAFC;font-size: 20pt;font-weight: 700;font-style: italic">First</span><span style="color: #F97316;font-size: 20pt"> pair</span></p>' +
        '<p style="text-align: center"><span style="color: #38BDF8">Center</span></p>' +
        '<p style="text-align: right"><span>Last</span></p>',
    );
    element.width = 200;
    element.height = 100;
    const input = document();
    input.slides = [{ ...slide(''), elements: [element] }];

    const source = UTF8_DECODER.decode(
      renderPptxDocumentToSvg(input).slides[0]?.data,
    );

    expect(source).toContain(
      '<text x="4" y="24" text-anchor="start" font-family="sans-serif"><tspan fill="#F8FAFC" font-size="20" font-weight="700" font-style="italic">First</tspan><tspan fill="#F97316" font-size="20"> pair</tspan></text>',
    );
    expect(source).toContain(
      '<text x="100" y="40" text-anchor="middle" font-family="sans-serif"><tspan fill="#38BDF8" font-size="12">Center</tspan></text>',
    );
    expect(source).not.toContain('Stryker was here');
    expect(source).toContain(
      '<text x="196" y="56" text-anchor="end" font-family="sans-serif"><tspan fill="#111827" font-size="12">Last</tspan></text>',
    );
  });

  it.each([
    ['mid', [46, 62, 78]],
    ['down', [64, 80, 96]],
    ['up', [24, 40, 56]],
  ])(
    'positions multiline text at vertical alignment %s',
    (vAlign, baselines) => {
      const element = text(
        '<p><span style="font-size: 20pt">First</span></p><p>Second</p><p>Third</p>',
      );
      element.width = 200;
      element.height = 100;
      element.vAlign = vAlign;
      const input = document();
      input.slides = [{ ...slide(''), elements: [element] }];

      const source = UTF8_DECODER.decode(
        renderPptxDocumentToSvg(input).slides[0]?.data,
      );

      expect(
        [...source.matchAll(/<text x="4" y="([\d.]+)"/g)].map((match) =>
          Number(match[1]),
        ),
      ).toEqual(baselines);
    },
  );

  it('clamps centered content to the top when it is taller than its box', () => {
    const element = text('<p>First</p><p>Second</p>');
    element.height = 20;
    element.vAlign = 'mid';
    const input = document();
    input.slides = [{ ...slide(''), elements: [element] }];

    const source = UTF8_DECODER.decode(
      renderPptxDocumentToSvg(input).slides[0]?.data,
    );

    expect(source).toContain('<text x="4" y="16"');
    expect(source).toContain('<text x="4" y="32"');
  });

  it('omits a text body for non-string or empty content', () => {
    const invalid = text('');
    invalid.content = undefined as never;
    const empty = text('<p><span></span></p>');
    const input = document();
    input.slides = [{ ...slide(''), elements: [invalid, empty] }];

    const source = UTF8_DECODER.decode(
      renderPptxDocumentToSvg(input).slides[0]?.data,
    );
    expect(source).not.toContain('<tspan');
    expect(source).not.toContain('overflow="hidden"');
    expect(source).not.toContain('Stryker was here');
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

describe('PowerPoint package SVG rendering', () => {
  it.each(PACKAGE_INPUTS)(
    'opens a %s package and renders it without an Office runtime',
    async (_name, input) => {
      const bytes = await createMinimalPptx();
      const result = await renderPptxToSvg(input(bytes), { scale: 2 });

      expect(result.slides).toHaveLength(1);
      expect(result.slides[0]).toMatchObject({
        format: 'svg',
        height: 810,
        mimeType: 'image/svg+xml',
        slideNumber: 1,
        width: 1440,
      });
      const source = UTF8_DECODER.decode(result.slides[0]?.data);
      expect(source).toContain('<title>PowerPoint slide 1</title>');
      expect(source).toContain('Hello AI');
    },
  );

  it('applies package parse limits before producing render output', async () => {
    const bytes = await createMinimalPptx();

    await expect(
      renderPptxToSvg(bytes, {
        parseLimits: { maxInputBytes: bytes.byteLength - 1 },
      }),
    ).rejects.toThrow(
      `PPTX resource limit maxInputBytes exceeded: ${bytes.byteLength} > ${bytes.byteLength - 1}`,
    );
  });

  it('applies slide selection after parsing the package', async () => {
    const bytes = await createMinimalPptx();

    await expect(renderPptxToSvg(bytes, { slideNumbers: [] })).resolves.toEqual(
      { slides: [] },
    );
  });
});

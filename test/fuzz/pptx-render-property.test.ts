import fc from 'fast-check';
import { SaxesParser } from 'saxes';
import { describe, expect, it } from 'vitest';

import {
  renderPptxDocumentToSvg,
  type PptxDocument,
  type PptxElement,
} from '../../src';
import type {
  Audio,
  CommonChart,
  Fill,
  Group,
  Image,
  Math as MathElement,
  Shape,
  Table,
  Text,
  Video,
} from '../../src/formats/pptx';
import { renderPptxDocumentToPng } from '../../src/formats/pptx/node';

const SVG_FUZZ_SEED = 0x53_56_47;
const PNG_FUZZ_SEED = 0x50_4e_47;
const SVG_FUZZ_RUNS = 96;
const PNG_FUZZ_RUNS = 24;
const UTF8_DECODER = new TextDecoder();
const PNG_SIGNATURE = [137, 80, 78, 71, 13, 10, 26, 10];
const SAFE_PNG_DATA_URI =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const hostileString = fc.oneof(
  fc.string({ maxLength: 80 }),
  fc
    .array(fc.integer({ min: 0, max: 31 }), {
      maxLength: 12,
      minLength: 1,
    })
    .map((points) => String.fromCharCode(...points)),
  fc
    .array(fc.integer({ min: 0xd800, max: 0xdfff }), {
      maxLength: 6,
      minLength: 1,
    })
    .map((points) => String.fromCharCode(...points)),
  fc.constantFrom(
    '<script>alert(1)</script>',
    '<foreignObject><iframe src="https://attacker.invalid"></iframe></foreignObject>',
    '"><image href="file:///etc/passwd"/>',
    '&lt;script&gt;alert(1)&lt;/script&gt;',
    'javascript:alert(1)',
    'https://attacker.invalid/payload',
    'blob:https://attacker.invalid/id',
    'file:///etc/passwd',
    'A & B < C > D "quoted" \'single\'',
  ),
);

const coordinate = fc.double({
  max: 900,
  min: -200,
  noDefaultInfinity: true,
  noNaN: true,
});
const extent = fc.double({
  max: 400,
  min: 0.01,
  noDefaultInfinity: true,
  noNaN: true,
});
const rotation = fc.double({
  max: 720,
  min: -720,
  noDefaultInfinity: true,
  noNaN: true,
});
const borderType = fc.constantFrom<'solid' | 'dashed' | 'dotted'>(
  'solid',
  'dashed',
  'dotted',
);
const color = fc.oneof(
  fc.constantFrom('#000000', '#ffffff', '#123abc', '#abcdef80'),
  hostileString,
);
const fill: fc.Arbitrary<Fill | null> = fc.oneof(
  fc.constant(null),
  color.map((value): Fill => ({ type: 'color', value })),
);
const box = fc.record({
  height: extent,
  left: coordinate,
  top: coordinate,
  width: extent,
});

function textArbitrary(): fc.Arbitrary<Text> {
  return fc
    .record({
      borderColor: color,
      borderStrokeDasharray: hostileString,
      borderType,
      borderWidth: fc.integer({ max: 8, min: 0 }),
      box,
      content: hostileString,
      fill,
      id: hostileString,
      isFlipH: fc.boolean(),
      isFlipV: fc.boolean(),
      isVertical: fc.boolean(),
      link: hostileString,
      name: hostileString,
      order: fc.integer({ max: 100, min: 0 }),
      rotate: rotation,
    })
    .map((value): Text => ({
      ...value.box,
      borderColor: value.borderColor,
      borderStrokeDasharray: value.borderStrokeDasharray,
      borderType: value.borderType,
      borderWidth: value.borderWidth,
      content: value.content,
      fill: value.fill,
      id: value.id,
      isFlipH: value.isFlipH,
      isFlipV: value.isFlipV,
      isVertical: value.isVertical,
      link: value.link,
      name: value.name,
      order: value.order,
      rotate: value.rotate,
      type: 'text',
      vAlign: 'top',
      wrap: true,
    }));
}

function shapeArbitrary(): fc.Arbitrary<Shape> {
  return fc
    .record({
      borderColor: color,
      borderStrokeDasharray: hostileString,
      borderType,
      borderWidth: fc.integer({ max: 8, min: 0 }),
      box,
      content: hostileString,
      fill,
      id: hostileString,
      isFlipH: fc.boolean(),
      isFlipV: fc.boolean(),
      link: hostileString,
      name: hostileString,
      order: fc.integer({ max: 100, min: 0 }),
      path: hostileString,
      rotate: rotation,
      shapType: fc.constantFrom(
        'rect',
        'roundRect',
        'ellipse',
        'line',
        'custom',
      ),
    })
    .map((value): Shape => ({
      ...value.box,
      borderColor: value.borderColor,
      borderStrokeDasharray: value.borderStrokeDasharray,
      borderType: value.borderType,
      borderWidth: value.borderWidth,
      content: value.content,
      fill: value.fill,
      id: value.id,
      isFlipH: value.isFlipH,
      isFlipV: value.isFlipV,
      link: value.link,
      name: value.name,
      order: value.order,
      path: value.path,
      pathViewBox: { height: 100, width: 100, x: 0, y: 0 },
      rotate: value.rotate,
      shapType: value.shapType,
      type: 'shape',
      vAlign: 'top',
      wrap: true,
    }));
}

function imageArbitrary(): fc.Arbitrary<Image> {
  const source = fc.oneof(
    fc.constant(SAFE_PNG_DATA_URI),
    hostileString,
    fc.constantFrom(
      'data:image/svg+xml;base64,PHN2Zy8+',
      'data:text/html;base64,PHNjcmlwdD5hbGVydCgxKTwvc2NyaXB0Pg==',
      'data:image/png;base64,not-canonical!',
    ),
  );
  return fc
    .record({
      base64: source,
      blob: hostileString,
      borderColor: color,
      borderStrokeDasharray: hostileString,
      borderType,
      borderWidth: fc.integer({ max: 8, min: 0 }),
      box,
      id: hostileString,
      isFlipH: fc.boolean(),
      isFlipV: fc.boolean(),
      link: hostileString,
      order: fc.integer({ max: 100, min: 0 }),
      ref: hostileString,
      rotate: rotation,
    })
    .map((value): Image => ({
      ...value.box,
      base64: value.base64,
      blob: value.blob,
      borderColor: value.borderColor,
      borderStrokeDasharray: value.borderStrokeDasharray,
      borderType: value.borderType,
      borderWidth: value.borderWidth,
      geom: 'rect',
      id: value.id,
      isFlipH: value.isFlipH,
      isFlipV: value.isFlipV,
      link: value.link,
      order: value.order,
      ref: value.ref,
      rotate: value.rotate,
      type: 'image',
    }));
}

function tableArbitrary(): fc.Arbitrary<Table> {
  return fc
    .record({
      box,
      id: hostileString,
      order: fc.integer({ max: 100, min: 0 }),
      text: hostileString,
    })
    .map((value): Table => ({
      ...value.box,
      borders: {},
      colWidths: [value.box.width],
      data: [
        [
          {
            borders: {},
            text: value.text,
            vAlign: 'top',
          },
        ],
      ],
      id: value.id,
      order: value.order,
      rowHeights: [value.box.height],
      type: 'table',
    }));
}

function chartArbitrary(): fc.Arbitrary<CommonChart> {
  return fc
    .record({
      box,
      color,
      id: hostileString,
      key: hostileString,
      order: fc.integer({ max: 100, min: 0 }),
      x: hostileString,
      y: fc.double({
        max: 1_000_000,
        min: -1_000_000,
        noDefaultInfinity: true,
        noNaN: true,
      }),
    })
    .map((value): CommonChart => ({
      ...value.box,
      chartType: 'barChart',
      colors: [value.color],
      data: [
        {
          key: value.key,
          values: [{ x: value.x, y: value.y }],
          xlabels: {},
        },
      ],
      id: value.id,
      order: value.order,
      type: 'chart',
    }));
}

function passiveMediaArbitrary(): fc.Arbitrary<Audio | Video> {
  return fc
    .record({
      blob: hostileString,
      box,
      id: hostileString,
      order: fc.integer({ max: 100, min: 0 }),
      ref: hostileString,
      rotate: rotation,
      type: fc.constantFrom<'audio' | 'video'>('audio', 'video'),
    })
    .map((value): Audio | Video => ({ ...value.box, ...value }));
}

function mathArbitrary(): fc.Arbitrary<MathElement> {
  return fc
    .record({
      box,
      id: hostileString,
      latex: hostileString,
      order: fc.integer({ max: 100, min: 0 }),
      picBase64: hostileString,
      picBlob: hostileString,
      picRef: hostileString,
      text: hostileString,
    })
    .map((value): MathElement => ({ ...value.box, ...value, type: 'math' }));
}

const leafElement: fc.Arbitrary<PptxElement> = fc.oneof(
  textArbitrary(),
  shapeArbitrary(),
  imageArbitrary(),
  tableArbitrary(),
  chartArbitrary(),
  passiveMediaArbitrary(),
  mathArbitrary(),
);

const element: fc.Arbitrary<PptxElement> = fc.oneof(
  leafElement,
  fc
    .record({
      box,
      elements: fc.array(leafElement, { maxLength: 3 }),
      id: hostileString,
      isFlipH: fc.boolean(),
      isFlipV: fc.boolean(),
      order: fc.integer({ max: 100, min: 0 }),
      rotate: rotation,
    })
    .map((value): Group => ({
      ...value.box,
      elements: value.elements,
      id: value.id,
      isFlipH: value.isFlipH,
      isFlipV: value.isFlipV,
      order: value.order,
      rotate: value.rotate,
      type: 'group',
    })),
);

function document(elements: PptxElement[]): PptxDocument {
  return {
    size: { height: 405, width: 720 },
    slides: [
      {
        elements,
        fill: { type: 'color', value: '#ffffff' },
        layoutElements: [],
        note: 'Untrusted note must never become SVG markup',
      },
    ],
    themeColors: [],
    usedFonts: [],
  };
}

function expectWellFormedXml(source: string): void {
  let invalid = false;
  const parser = new SaxesParser({ xmlns: true });
  parser.on('error', () => {
    invalid = true;
  });
  parser.write(source).close();
  expect(invalid).toBe(false);
}

function expectSelfContainedSafeSvg(source: string): void {
  expectWellFormedXml(source);
  expect(source).not.toMatch(/<(?:script|foreignObject)\b/i);
  expect(source).not.toMatch(/\b(?:NaN|Infinity)\b/);
  expect(source).not.toMatch(
    /\b(?:href|src|xlink:href)="(?!data:image\/(?:gif|jpeg|png|webp);base64,)/i,
  );
}

function pngDimensions(data: Uint8Array): { height: number; width: number } {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  return { height: view.getUint32(20), width: view.getUint32(16) };
}

describe('PowerPoint render seeded properties', () => {
  it('renders arbitrary untrusted scenes deterministically without external references or caller mutation', () => {
    fc.assert(
      fc.property(fc.array(element, { maxLength: 12 }), (elements) => {
        const input = document(elements);
        const before = structuredClone(input);

        const first = renderPptxDocumentToSvg(input, { scale: 0.5 });
        const second = renderPptxDocumentToSvg(input, { scale: 0.5 });

        expect(input).toEqual(before);
        expect(first).toEqual(second);
        expect(first.slides).toHaveLength(1);
        const slide = first.slides[0];
        if (slide === undefined) throw new Error('Expected one rendered slide');
        expect(slide).toMatchObject({
          format: 'svg',
          height: 203,
          mimeType: 'image/svg+xml',
          slideNumber: 1,
          width: 360,
        });
        expectSelfContainedSafeSvg(UTF8_DECODER.decode(slide.data));
      }),
      { endOnFailure: true, numRuns: SVG_FUZZ_RUNS, seed: SVG_FUZZ_SEED },
    );
  });

  it('skips arbitrary invalid element geometry without leaking non-finite values', () => {
    const invalidExtent = fc.constantFrom(
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.NEGATIVE_INFINITY,
      0,
      -1,
    );
    fc.assert(
      fc.property(
        textArbitrary(),
        invalidExtent,
        fc.constantFrom<'height' | 'width'>('height', 'width'),
        (text, invalid, property) => {
          text[property] = invalid;
          const slide = renderPptxDocumentToSvg(document([text])).slides[0];
          if (slide === undefined)
            throw new Error('Expected one rendered slide');
          const source = UTF8_DECODER.decode(slide.data);

          expectSelfContainedSafeSvg(source);
          expect(source).not.toContain('<tspan');
          expect(slide.warnings).toContainEqual({
            code: 'approximate-shape',
            ...(text.id === '' ? {} : { elementId: text.id }),
            message: 'The preview skipped an element with invalid geometry.',
            slideNumber: 1,
          });
        },
      ),
      { endOnFailure: true, numRuns: SVG_FUZZ_RUNS, seed: SVG_FUZZ_SEED + 1 },
    );
  });

  it('rasterizes arbitrary portable text to PNG with independently readable dimensions', () => {
    fc.assert(
      fc.property(
        fc.array(fc.oneof(textArbitrary(), tableArbitrary()), {
          maxLength: 8,
        }),
        (elements) => {
          const input = document(elements);
          const before = structuredClone(input);
          const result = renderPptxDocumentToPng(input, { scale: 0.25 });
          const slide = result.slides[0];

          expect(input).toEqual(before);
          expect(slide).toMatchObject({
            format: 'png',
            height: 101,
            mimeType: 'image/png',
            slideNumber: 1,
            width: 180,
          });
          expect(Array.from(slide?.data.slice(0, 8) ?? [])).toEqual(
            PNG_SIGNATURE,
          );
          expect(slide ? pngDimensions(slide.data) : null).toEqual({
            height: 101,
            width: 180,
          });
        },
      ),
      { endOnFailure: true, numRuns: PNG_FUZZ_RUNS, seed: PNG_FUZZ_SEED },
    );
  }, 30_000);
});

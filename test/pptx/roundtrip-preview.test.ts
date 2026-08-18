import { describe, expect, it } from 'vitest';

import { createPptx } from '../../src/formats/pptx/creator';
import { parse } from '../../src/formats/pptx/parser';
import type { PptxDocument } from '../../src/formats/pptx/types';
import type { PptxSceneDocument } from '../../src/formats/pptx/scene-types';
import { validatePptxScene } from '../../src/formats/pptx/scene-validation';
import {
  createPowerPointRoundTripPreview,
  plainTextFromPowerPointHtml,
} from '../../src/formats/pptx/roundtrip/preview';

function sourceScene(): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {
              transform: {
                flipHorizontal: true,
                flipVertical: false,
                height: 80,
                rotation: 15,
                width: 300,
                x: 20,
                y: 30,
              },
            },
            key: 'source-text',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    { key: 'source-run', text: 'Preview text', type: 'run' },
                  ],
                  key: 'source-paragraph',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'source-slide',
      },
      { elements: [], key: 'empty-slide' },
    ],
    themes: [],
  };
}

describe('PowerPoint round-trip semantic preview', () => {
  it('creates a deterministic preservation-only scene from parsed output', async () => {
    const source = await createPptx(sourceScene());
    const parsed = await parse(source.data, {
      errorMode: 'strict',
      imageMode: 'none',
    });

    const first = createPowerPointRoundTripPreview(parsed);
    const second = createPowerPointRoundTripPreview(parsed);

    expect(second).toEqual(first);
    expect(validatePptxScene(first)).toEqual({ issues: [], valid: true });
    expect(first.size).toEqual({ height: 540, width: 960 });
    expect(first.slides.map((slide) => slide.key)).toEqual([
      'slide-1',
      'slide-2',
    ]);
    expect(first.slides[1]?.elements).toEqual([]);
    expect(first.slides[0]?.elements).toEqual([
      {
        authored: {},
        key: 'slide-1-element-1',
        name: 'Text Box 2',
        resolved: {
          hidden: false,
          transform: {
            flipHorizontal: true,
            flipVertical: false,
            height: 80,
            rotation: 15,
            width: 300,
            x: 20,
            y: 30,
          },
        },
        text: {
          body: {
            anchor: 'top',
            vertical: false,
            wrap: true,
          },
          paragraphs: [
            {
              children: [
                {
                  key: 'slide-1-element-1-run-1',
                  text: 'Preview text',
                  type: 'run',
                },
              ],
              key: 'slide-1-element-1-paragraph-1',
            },
          ],
        },
        type: 'text',
      },
    ]);
  });

  it('decodes portable PowerPoint HTML into ordered plain text', () => {
    expect(
      plainTextFromPowerPointHtml(
        '<p><span>A&nbsp;&lt;&amp;</span><br><span>B</span></p><p>C&#x21;</p>',
      ),
    ).toBe('A <&\nB\nC!');
    expect(plainTextFromPowerPointHtml('<p>A</p><p></p><p></p>')).toBe('A');
  });

  it.each([
    ['down', 'bottom'],
    ['mid', 'center'],
    ['dist', 'distributed'],
    ['just', 'justified'],
    ['unknown', 'top'],
  ])('maps vertical alignment %s to %s', async (vAlign, anchor) => {
    const source = await createPptx(sourceScene());
    const parsed = await parse(source.data, { imageMode: 'none' });
    const element = parsed.slides[0]?.elements[0];
    if (element?.type !== 'text') throw new Error('Expected text element');
    element.vAlign = vAlign;

    const preview = createPowerPointRoundTripPreview(parsed);

    expect(preview.slides[0]?.elements[0]).toMatchObject({
      text: { body: { anchor } },
    });
  });

  it('includes a defined autofit mode and omits an absent one', async () => {
    const source = await createPptx(sourceScene());
    const parsed = await parse(source.data, { imageMode: 'none' });
    const element = parsed.slides[0]?.elements[0];
    if (element?.type !== 'text') throw new Error('Expected text element');
    element.autoFit = { type: 'shape' };
    expect(
      createPowerPointRoundTripPreview(parsed).slides[0]?.elements[0],
    ).toMatchObject({ text: { body: { autoFit: 'shape' } } });
    delete element.autoFit;
    const previewElement =
      createPowerPointRoundTripPreview(parsed).slides[0]?.elements[0];
    if (previewElement?.type !== 'text')
      throw new Error('Expected preview text');
    expect(previewElement.text.body).not.toHaveProperty('autoFit');
  });

  it('maps empty native shapes and preserves text-bearing shapes as opaque', () => {
    const nativeShape = {
      content: '',
      height: 25,
      isFlipH: true,
      isFlipV: false,
      left: 5,
      name: 'Native shape',
      rotate: 10,
      top: 6,
      type: 'shape',
      width: 35,
    };
    const unsupportedWithText = {
      content: 'Fallback',
      height: 40,
      isFlipH: false,
      isFlipV: true,
      left: 10,
      name: 'Shape',
      rotate: 5,
      top: 20,
      type: 'shape',
      width: 30,
    };
    const unsupportedWithoutText = {
      height: 20,
      left: 1,
      name: 'Broken',
      top: 2,
      type: 'image',
      width: 30,
    };
    const document = {
      size: { height: 540, width: 960 },
      slides: [
        {
          elements: [],
          fill: { type: 'color', value: '#ffffff' },
          layoutElements: [],
          note: '',
        },
        {
          elements: [nativeShape, unsupportedWithText, unsupportedWithoutText],
          fill: { type: 'color', value: '#ffffff' },
          layoutElements: [],
          note: '',
        },
      ],
      themeColors: [],
      usedFonts: [],
    } as unknown as PptxDocument;

    const preview = createPowerPointRoundTripPreview(document);

    expect(preview.slides[1]?.elements).toEqual([
      {
        authored: {},
        key: 'slide-2-element-1',
        name: 'Native shape',
        resolved: {
          hidden: false,
          transform: {
            flipHorizontal: true,
            flipVertical: false,
            height: 25,
            rotation: 10,
            width: 35,
            x: 5,
            y: 6,
          },
        },
        type: 'shape',
      },
      {
        authored: {},
        feature: 'shape',
        key: 'slide-2-element-2',
        previewText: 'Fallback',
        resolved: {
          hidden: false,
          transform: {
            flipHorizontal: false,
            flipVertical: true,
            height: 40,
            rotation: 5,
            width: 30,
            x: 10,
            y: 20,
          },
        },
        type: 'unsupported',
      },
      {
        authored: {},
        key: 'slide-2-element-3',
        resolved: {
          hidden: false,
          transform: { height: 20, width: 30, x: 1, y: 2 },
        },
        type: 'image',
      },
    ]);
  });

  it('does not mutate the parsed document', async () => {
    const source = await createPptx(sourceScene());
    const parsed = await parse(source.data, { imageMode: 'none' });
    const before = structuredClone(parsed);

    createPowerPointRoundTripPreview(parsed);

    expect(parsed).toEqual(before);
  });

  it.each([
    ['left', Number.NaN],
    ['top', Number.POSITIVE_INFINITY],
    ['width', Number.NaN],
    ['width', 0],
    ['height', Number.NEGATIVE_INFINITY],
    ['height', 0],
  ] as const)(
    'omits a transform with non-rendering %s %s from the portable preview',
    async (property, value) => {
      const source = await createPptx(sourceScene());
      const parsed = await parse(source.data, { imageMode: 'none' });
      const element = parsed.slides[0]?.elements[0];
      if (!element) throw new Error('Expected a parsed preview element');
      element[property] = value;

      const preview = createPowerPointRoundTripPreview(parsed);

      expect(preview.slides[0]?.elements[0]?.resolved).toEqual({
        hidden: false,
      });
      expect(validatePptxScene(preview)).toEqual({ issues: [], valid: true });
    },
  );
});

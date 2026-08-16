import { describe, expect, it } from 'vitest';

import { createPptx } from '../../src/formats/pptx/creator';
import { parse } from '../../src/formats/pptx/parser';
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

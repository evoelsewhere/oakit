import { describe, expect, it } from 'vitest';

import { createPptx } from '../../src/formats/pptx/creator';
import { parse } from '../../src/formats/pptx/parser';
import type { PptxSceneDocument } from '../../src/formats/pptx/scene-types';
import { validatePptxScene } from '../../src/formats/pptx/scene-validation';
import { createPowerPointRoundTripPreview } from '../../src/formats/pptx/roundtrip/preview';

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
        feature: 'text',
        key: 'slide-1-element-1',
        previewText:
          '<p style="text-align: left;"><span style="font-size: 18pt;font-family: &quot;Aptos&quot;;">Preview&nbsp;text</span></p>',
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
        type: 'unsupported',
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
});

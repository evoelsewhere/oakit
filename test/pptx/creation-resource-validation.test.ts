import { describe, expect, it } from 'vitest';

import {
  MAX_POWERPOINT_CREATION_ELEMENTS,
  MAX_POWERPOINT_CREATION_MEDIA,
  MAX_POWERPOINT_CREATION_MEDIA_BYTES,
  MAX_POWERPOINT_CREATION_PARAGRAPHS,
  MAX_POWERPOINT_CREATION_STRING_CODE_UNITS,
  MAX_POWERPOINT_CREATION_TOTAL_CHART_POINTS,
  MAX_POWERPOINT_CREATION_TEXT_NODES,
  MAX_POWERPOINT_CREATION_TOTAL_MEDIA_BYTES,
} from '../../src/formats/pptx/creation-limits';
import { validatePowerPointCreationResources } from '../../src/formats/pptx/creation-resource-validation';

function document(slides: unknown[]): Record<string, unknown> {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides,
    themes: [],
  };
}

describe('PowerPoint creation resource validation', () => {
  it('accepts resource counts exactly at every boundary', () => {
    const input = document([
      {
        elements: new Array(MAX_POWERPOINT_CREATION_ELEMENTS),
        key: 'slide-1',
      },
    ]);

    expect(validatePowerPointCreationResources(input)).toEqual([]);
  });

  it('accepts native media budgets exactly at every boundary', () => {
    const exactCount = document([]);
    exactCount.media = new Array(MAX_POWERPOINT_CREATION_MEDIA);
    const exactPart = document([]);
    exactPart.media = [
      { data: { byteLength: MAX_POWERPOINT_CREATION_MEDIA_BYTES } },
    ];
    const exactTotal = document([]);
    exactTotal.media = Array.from({ length: 4 }, () => ({
      data: { byteLength: MAX_POWERPOINT_CREATION_TOTAL_MEDIA_BYTES / 4 },
    }));

    for (const input of [exactCount, exactPart, exactTotal]) {
      expect(
        validatePowerPointCreationResources(input, 'create-native-v1'),
      ).toEqual([]);
    }
  });

  it('accepts paragraph, text-node, and string counts at their boundaries', () => {
    const paragraphInput = document([
      {
        elements: [
          {
            type: 'text',
            text: {
              paragraphs: new Array(MAX_POWERPOINT_CREATION_PARAGRAPHS),
            },
          },
        ],
      },
    ]);
    const textNodeInput = document([
      {
        elements: [
          {
            type: 'text',
            text: {
              paragraphs: [
                {
                  children: new Array(MAX_POWERPOINT_CREATION_TEXT_NODES),
                },
              ],
            },
          },
        ],
      },
    ]);
    const stringInput = {
      slides: [],
      value: 'a'.repeat(MAX_POWERPOINT_CREATION_STRING_CODE_UNITS),
    };

    expect(validatePowerPointCreationResources(paragraphInput)).toEqual([]);
    expect(validatePowerPointCreationResources(textNodeInput)).toEqual([]);
    expect(validatePowerPointCreationResources(stringInput)).toEqual([]);
  });

  it('rejects an element count beyond the package budget', () => {
    const input = document([
      {
        elements: new Array(MAX_POWERPOINT_CREATION_ELEMENTS + 1),
        key: 'slide-1',
      },
    ]);

    expect(validatePowerPointCreationResources(input)).toEqual([
      {
        code: 'resource-limit-exceeded',
        message: `Creation profile create-text-v1 supports at most ${MAX_POWERPOINT_CREATION_ELEMENTS} elements`,
        path: '$.slides',
      },
    ]);
  });

  it('counts every element nested inside native groups', () => {
    const input = document([
      {
        elements: [
          {
            elements: new Array(MAX_POWERPOINT_CREATION_ELEMENTS),
            type: 'group',
          },
        ],
        key: 'slide-1',
      },
    ]);

    expect(
      validatePowerPointCreationResources(input, 'create-native-v1'),
    ).toContainEqual({
      code: 'resource-limit-exceeded',
      message: `Creation profile create-native-v1 supports at most ${MAX_POWERPOINT_CREATION_ELEMENTS} elements`,
      path: '$.slides',
    });
  });

  it('enforces the aggregate chart-point budget across native charts', () => {
    const exact = document([
      {
        elements: [
          {
            series: [
              {
                categories: new Array<string>(
                  MAX_POWERPOINT_CREATION_TOTAL_CHART_POINTS,
                ),
                values: new Array<number>(
                  MAX_POWERPOINT_CREATION_TOTAL_CHART_POINTS,
                ),
              },
            ],
            type: 'chart',
          },
        ],
      },
    ]);
    expect(
      validatePowerPointCreationResources(exact, 'create-native-v1'),
    ).toEqual([]);

    const pointsPerSeries = 10_000;
    const seriesCount =
      Math.floor(MAX_POWERPOINT_CREATION_TOTAL_CHART_POINTS / pointsPerSeries) +
      1;
    const input = document([
      {
        elements: [
          {
            series: Array.from({ length: seriesCount }, () => ({
              categories: new Array<string>(pointsPerSeries),
              values: new Array<number>(pointsPerSeries),
            })),
            type: 'chart',
          },
        ],
      },
    ]);

    expect(
      validatePowerPointCreationResources(input, 'create-native-v1'),
    ).toContainEqual({
      code: 'resource-limit-exceeded',
      message: `Creation profile create-native-v1 supports at most ${MAX_POWERPOINT_CREATION_TOTAL_CHART_POINTS} total chart points`,
      path: '$.slides',
    });
  });

  it('ignores primitive entries while traversing nested group resources', () => {
    const input = document([
      {
        elements: [
          null,
          'primitive',
          { elements: [undefined, 7, false], type: 'group' },
        ],
        key: 'slide-1',
      },
    ]);

    expect(
      validatePowerPointCreationResources(input, 'create-native-v1'),
    ).toEqual([]);
  });

  it('rejects a paragraph count beyond the XML-node budget', () => {
    const input = document([
      {
        elements: [
          {
            type: 'text',
            text: {
              paragraphs: new Array(MAX_POWERPOINT_CREATION_PARAGRAPHS + 1),
            },
          },
        ],
      },
    ]);

    expect(validatePowerPointCreationResources(input)).toEqual([
      {
        code: 'resource-limit-exceeded',
        message: `Creation profile create-text-v1 supports at most ${MAX_POWERPOINT_CREATION_PARAGRAPHS} paragraphs`,
        path: '$.slides',
      },
    ]);
  });

  it('rejects a text-node count beyond the XML-node budget', () => {
    const input = document([
      {
        elements: [
          {
            type: 'text',
            text: {
              paragraphs: [
                {
                  children: new Array(MAX_POWERPOINT_CREATION_TEXT_NODES + 1),
                },
              ],
            },
          },
        ],
      },
    ]);

    expect(validatePowerPointCreationResources(input)).toEqual([
      {
        code: 'resource-limit-exceeded',
        message: `Creation profile create-text-v1 supports at most ${MAX_POWERPOINT_CREATION_TEXT_NODES} text nodes`,
        path: '$.slides',
      },
    ]);
  });

  it('counts strings recursively across objects and arrays', () => {
    const input = document([]);
    input.extra = [
      'a'.repeat(MAX_POWERPOINT_CREATION_STRING_CODE_UNITS / 2),
      { nested: 'b'.repeat(MAX_POWERPOINT_CREATION_STRING_CODE_UNITS / 2 + 1) },
    ];

    expect(validatePowerPointCreationResources(input)).toEqual([
      {
        code: 'resource-limit-exceeded',
        message: `Creation profile create-text-v1 supports at most ${MAX_POWERPOINT_CREATION_STRING_CODE_UNITS} string code units`,
        path: '$',
      },
    ]);
  });

  it('ignores non-string primitive values while measuring strings', () => {
    const input = document([]);
    input.values = [null, undefined, false, true, 0, 1];

    expect(validatePowerPointCreationResources(input)).toEqual([]);
  });

  it('counts native shapes without traversing nonexistent text', () => {
    const input = document([
      {
        elements: [
          { key: 'shape-1', type: 'shape' },
          {
            key: 'text-1',
            type: 'text',
            text: { paragraphs: [{ children: [{ text: 'value' }] }] },
          },
        ],
        key: 'slide-1',
      },
    ]);

    expect(
      validatePowerPointCreationResources(input, 'create-native-v1'),
    ).toEqual([]);
  });

  it('counts structured text inside every native table cell', () => {
    const input = document([
      {
        elements: [
          {
            rows: [
              {
                cells: Array.from(
                  { length: MAX_POWERPOINT_CREATION_PARAGRAPHS + 1 },
                  () => ({
                    text: { paragraphs: [{ children: [] }] },
                  }),
                ),
              },
            ],
            type: 'table',
          },
        ],
      },
    ]);

    expect(
      validatePowerPointCreationResources(input, 'create-native-v1'),
    ).toContainEqual({
      code: 'resource-limit-exceeded',
      message: `Creation profile create-native-v1 supports at most ${MAX_POWERPOINT_CREATION_PARAGRAPHS} paragraphs`,
      path: '$.slides',
    });
  });

  it('enforces media count, per-resource, and aggregate byte budgets', () => {
    const tooMany = document([]);
    tooMany.media = new Array(MAX_POWERPOINT_CREATION_MEDIA + 1);
    const tooLarge = document([]);
    tooLarge.media = [
      { data: { byteLength: MAX_POWERPOINT_CREATION_MEDIA_BYTES + 1 } },
    ];
    const tooLargeTogether = document([]);
    tooLargeTogether.media = [
      {
        data: { byteLength: MAX_POWERPOINT_CREATION_TOTAL_MEDIA_BYTES / 2 + 1 },
      },
      { data: { byteLength: MAX_POWERPOINT_CREATION_TOTAL_MEDIA_BYTES / 2 } },
    ];

    expect(
      validatePowerPointCreationResources(tooMany, 'create-native-v1'),
    ).toContainEqual({
      code: 'resource-limit-exceeded',
      message: `Creation profile create-native-v1 supports at most ${MAX_POWERPOINT_CREATION_MEDIA} media resources`,
      path: '$.media',
    });
    expect(
      validatePowerPointCreationResources(tooLarge, 'create-native-v1'),
    ).toContainEqual({
      code: 'resource-limit-exceeded',
      message: `Creation profile create-native-v1 supports at most ${MAX_POWERPOINT_CREATION_MEDIA_BYTES} bytes per media resource`,
      path: '$.media',
    });
    expect(
      validatePowerPointCreationResources(tooLargeTogether, 'create-native-v1'),
    ).toContainEqual({
      code: 'resource-limit-exceeded',
      message: `Creation profile create-native-v1 supports at most ${MAX_POWERPOINT_CREATION_TOTAL_MEDIA_BYTES} total media bytes`,
      path: '$.media',
    });
  });

  it('does not traverse binary media while counting string code units', () => {
    const input = document([]);
    const data = new Uint8Array(1024);
    Object.defineProperty(data, 'metadata', {
      enumerable: true,
      value: 'x'.repeat(MAX_POWERPOINT_CREATION_STRING_CODE_UNITS + 1),
    });
    input.media = [
      {
        data,
        key: 'a',
        mimeType: 'image/png',
      },
    ];

    expect(
      validatePowerPointCreationResources(input, 'create-native-v1'),
    ).toEqual([]);
  });
});

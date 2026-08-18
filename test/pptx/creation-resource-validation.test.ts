import { describe, expect, it } from 'vitest';

import {
  MAX_POWERPOINT_CREATION_ELEMENTS,
  MAX_POWERPOINT_CREATION_PARAGRAPHS,
  MAX_POWERPOINT_CREATION_STRING_CODE_UNITS,
  MAX_POWERPOINT_CREATION_TEXT_NODES,
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
});

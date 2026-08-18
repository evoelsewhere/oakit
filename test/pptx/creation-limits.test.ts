import { describe, expect, it } from 'vitest';

import {
  isSupportedPowerPointCreationSlideCount,
  MAX_POWERPOINT_CREATION_ELEMENTS,
  MAX_POWERPOINT_CREATION_MEDIA,
  MAX_POWERPOINT_CREATION_MEDIA_BYTES,
  MAX_POWERPOINT_CREATION_PARAGRAPHS,
  MAX_POWERPOINT_CREATION_SLIDES,
  MAX_POWERPOINT_CREATION_STRING_CODE_UNITS,
  MAX_POWERPOINT_CREATION_TEXT_NODES,
  MAX_POWERPOINT_CREATION_TOTAL_MEDIA_BYTES,
} from '../../src/formats/pptx/creation-limits';

describe('PowerPoint creation limits', () => {
  it('publishes the reviewed slide-count ceiling', () => {
    expect(MAX_POWERPOINT_CREATION_SLIDES).toBe(10_000);
  });

  it('publishes the reviewed scene resource ceilings', () => {
    expect({
      elements: MAX_POWERPOINT_CREATION_ELEMENTS,
      media: MAX_POWERPOINT_CREATION_MEDIA,
      mediaBytes: MAX_POWERPOINT_CREATION_MEDIA_BYTES,
      paragraphs: MAX_POWERPOINT_CREATION_PARAGRAPHS,
      stringCodeUnits: MAX_POWERPOINT_CREATION_STRING_CODE_UNITS,
      textNodes: MAX_POWERPOINT_CREATION_TEXT_NODES,
      totalMediaBytes: MAX_POWERPOINT_CREATION_TOTAL_MEDIA_BYTES,
    }).toEqual({
      elements: 5_000,
      media: 1_000,
      mediaBytes: 67_108_864,
      paragraphs: 10_000,
      stringCodeUnits: 8_388_608,
      textNodes: 40_000,
      totalMediaBytes: 268_435_456,
    });
  });

  it.each([0, 1, 413, 9_999, 10_000])(
    'accepts supported slide count %s',
    (slideCount) => {
      expect(isSupportedPowerPointCreationSlideCount(slideCount)).toBe(true);
    },
  );

  it.each([
    -1,
    10_001,
    1.5,
    Number.NaN,
    Number.POSITIVE_INFINITY,
    Number.MAX_SAFE_INTEGER + 1,
  ])('rejects unsupported slide count %s', (slideCount) => {
    expect(isSupportedPowerPointCreationSlideCount(slideCount)).toBe(false);
  });
});

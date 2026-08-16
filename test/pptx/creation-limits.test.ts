import { describe, expect, it } from 'vitest';

import {
  isSupportedPowerPointCreationSlideCount,
  MAX_POWERPOINT_CREATION_ELEMENTS,
  MAX_POWERPOINT_CREATION_PARAGRAPHS,
  MAX_POWERPOINT_CREATION_SLIDES,
  MAX_POWERPOINT_CREATION_STRING_CODE_UNITS,
  MAX_POWERPOINT_CREATION_TEXT_NODES,
} from '../../src/formats/pptx/creation-limits';

describe('PowerPoint creation limits', () => {
  it('publishes the reviewed slide-count ceiling', () => {
    expect(MAX_POWERPOINT_CREATION_SLIDES).toBe(10_000);
  });

  it('publishes the reviewed scene resource ceilings', () => {
    expect({
      elements: MAX_POWERPOINT_CREATION_ELEMENTS,
      paragraphs: MAX_POWERPOINT_CREATION_PARAGRAPHS,
      stringCodeUnits: MAX_POWERPOINT_CREATION_STRING_CODE_UNITS,
      textNodes: MAX_POWERPOINT_CREATION_TEXT_NODES,
    }).toEqual({
      elements: 5_000,
      paragraphs: 10_000,
      stringCodeUnits: 8_388_608,
      textNodes: 40_000,
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

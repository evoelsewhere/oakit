import { describe, expect, it } from 'vitest';

import {
  isSupportedPowerPointCreationSlideCount,
  MAX_POWERPOINT_CREATION_SLIDES,
} from '../../src/formats/pptx/creation-limits';

describe('PowerPoint creation limits', () => {
  it('publishes the reviewed slide-count ceiling', () => {
    expect(MAX_POWERPOINT_CREATION_SLIDES).toBe(10_000);
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

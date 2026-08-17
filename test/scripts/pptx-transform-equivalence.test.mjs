import { describe, expect, it } from 'vitest';

import {
  pptxTransformMatrix,
  pptxTransformsAreEquivalent,
} from '../../scripts/reliability/pptx-transform-equivalence.mjs';

describe('PowerPoint producer transform equivalence', () => {
  it('recognizes Google Slides double-flip rotation normalization', () => {
    expect(
      pptxTransformsAreEquivalent(
        { flipHorizontal: true, flipVertical: true, rotation: 45 },
        { flipHorizontal: false, flipVertical: false, rotation: -135 },
      ),
    ).toBe(true);
  });

  it('normalizes complete turns without hiding a visual reflection', () => {
    expect(
      pptxTransformsAreEquivalent(
        { flipHorizontal: false, flipVertical: false, rotation: 360 },
        { flipHorizontal: false, flipVertical: false, rotation: 0 },
      ),
    ).toBe(true);
    expect(
      pptxTransformsAreEquivalent(
        { flipHorizontal: true, flipVertical: false, rotation: 0 },
        { flipHorizontal: false, flipVertical: false, rotation: 0 },
      ),
    ).toBe(false);
  });

  it('returns the exact rotation and reflection matrix', () => {
    expect(
      pptxTransformMatrix({
        flipHorizontal: true,
        flipVertical: false,
        rotation: 90,
      }),
    ).toEqual([
      -Math.cos(Math.PI / 2),
      -Math.sin(Math.PI / 2),
      -Math.sin(Math.PI / 2),
      Math.cos(Math.PI / 2),
    ]);
  });

  it.each([
    [null, 'transform must be an object'],
    [
      { flipHorizontal: false, flipVertical: false, rotation: Number.NaN },
      'rotation must be finite',
    ],
    [
      { flipHorizontal: 0, flipVertical: false, rotation: 0 },
      'horizontal flip must be boolean',
    ],
  ])('rejects an invalid transform %#', (value, message) => {
    expect(() => pptxTransformMatrix(value)).toThrow(message);
  });
});

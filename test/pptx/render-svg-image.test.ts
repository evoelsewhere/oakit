import { describe, expect, it } from 'vitest';

import { svgImageCrop } from '../../src/formats/pptx/render-svg-image';

const BOX = { height: 80, left: 0, top: 0, width: 100 };

describe('PowerPoint SVG image crops', () => {
  it('maps canonical crop percentages into a clipped source viewport', () => {
    expect(svgImageCrop({ b: 0, l: 10, r: 20, t: 25 }, BOX)).toEqual({
      height: 106.66666666666667,
      width: 142.85714285714286,
      x: -14.285714285714286,
      y: -26.666666666666668,
    });
  });

  it('defaults omitted edges and preserves bounded negative expansion', () => {
    expect(svgImageCrop({ b: -20, l: 30 }, BOX)).toEqual({
      height: 66.66666666666667,
      width: 142.85714285714286,
      x: -42.857142857142854,
      y: 0,
    });
    expect(svgImageCrop({}, BOX)).toEqual({
      height: 80,
      width: 100,
      x: 0,
      y: 0,
    });
  });

  it('accepts exact crop bounds when the visible fraction stays positive', () => {
    expect(svgImageCrop({ l: -100 }, BOX)).toEqual({
      height: 80,
      width: 50,
      x: 50,
      y: 0,
    });
    expect(svgImageCrop({ l: 100, r: -1 }, BOX)).toEqual({
      height: 80,
      width: 9_999.99999999999,
      x: -9_999.99999999999,
      y: 0,
    });
    expect(svgImageCrop({ l: 101, r: -2 }, BOX)).toBeNull();
  });

  it.each([
    null,
    [],
    { l: Number.NaN },
    { r: Number.POSITIVE_INFINITY },
    { t: '1' },
    { b: -101 },
    { l: 101 },
    { l: 50, r: 50 },
    { b: 60, t: 40 },
  ])('rejects unsafe crop %#', (value) => {
    expect(svgImageCrop(value, BOX)).toBeNull();
  });
});

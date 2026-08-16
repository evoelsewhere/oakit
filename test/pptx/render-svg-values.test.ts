import { describe, expect, it } from 'vitest';

import {
  embeddedRasterDataUri,
  svgBox,
  svgColor,
  svgDashArray,
  svgNumber,
} from '../../src/formats/pptx/render-svg-values';

describe('PowerPoint SVG values', () => {
  it.each([
    [0, '0'],
    [-0, '0'],
    [1.234_54, '1.2345'],
    [1.234_56, '1.2346'],
    [-2.5, '-2.5'],
  ])('serializes finite number %s as %s', (value, expected) => {
    expect(svgNumber(value)).toBe(expected);
  });

  it.each([Number.NaN, Number.POSITIVE_INFINITY, '1', null])(
    'rejects unsafe SVG number %s',
    (value) => {
      expect(svgNumber(value)).toBeNull();
    },
  );

  it.each([
    ['#aabbcc', '#aabbcc'],
    ['AABBCCDD', '#AABBCCDD'],
    ['fff', null],
    ['red', null],
    [7, null],
  ])('normalizes SVG color %s', (value, expected) => {
    expect(svgColor(value)).toBe(expected);
  });

  it.each([
    ['1 2.34567,0', '1 2.3457 0'],
    [' 3 , 4 ', '3 4'],
    ['', null],
    ['   ', null],
    ['1 -2', null],
    ['1 nope', null],
    [null, null],
  ])('normalizes SVG dash array %s', (value, expected) => {
    expect(svgDashArray(value)).toBe(expected);
  });

  it.each(['png', 'jpeg', 'gif', 'webp'])(
    'accepts canonical embedded %s bytes and normalizes media case',
    (format) => {
      expect(
        embeddedRasterDataUri(`data:image/${format.toUpperCase()};base64,AA==`),
      ).toBe(`data:image/${format};base64,AA==`);
    },
  );

  it.each([
    'https://example.com/image.png',
    'file:///tmp/image.png',
    'data:image/svg+xml;base64,AA==',
    'data:text/html;base64,AA==',
    'data:image/png,AA==',
    'data:image/png;base64,',
    'data:image/png;base64,AB==',
    'data:image/png;base64,AA A=',
    'DATA:image/png;base64,AA==',
    'data:;base64,AA==',
  ])('rejects unsafe embedded image %s', (value) => {
    expect(embeddedRasterDataUri(value)).toBeNull();
  });

  it('rejects a non-string image value', () => {
    expect(embeddedRasterDataUri(new Uint8Array([0]))).toBeNull();
  });

  it('accepts finite boxes including negative positions', () => {
    expect(svgBox({ height: 20, left: -1, top: -2, width: 10 })).toEqual({
      height: 20,
      left: -1,
      top: -2,
      width: 10,
    });
  });

  it.each([
    null,
    7,
    {},
    { height: 1, left: 0, top: 0, width: 0 },
    { height: -1, left: 0, top: 0, width: 1 },
    { height: 0, left: 0, top: 0, width: 1 },
    { height: 1, left: Number.NaN, top: 0, width: 1 },
    { height: 1, left: 0, top: Number.POSITIVE_INFINITY, width: 1 },
    { height: 1, left: 0, top: 0, width: '1' },
    { height: '1', left: 0, top: 0, width: 1 },
  ])('rejects unsafe box %#', (value) => {
    expect(svgBox(value)).toBeNull();
  });
});

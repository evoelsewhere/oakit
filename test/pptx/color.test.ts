import { describe, expect, it } from 'vitest';

import {
  applyHueMod,
  applyLumMod,
  applyLumOff,
  applySatMod,
  applyShade,
  applyTint,
  getColorName2Hex,
  hslToRgb,
  hueToRgb,
} from '../../src/formats/pptx/internal/color';

describe('PPTX color-space conversion', () => {
  it.each([
    [-7, 0.2],
    [-1, 0.2],
    [0.5, 0.5],
    [1, 0.8],
    [2, 0.8],
    [3.5, 0.5],
    [5, 0.2],
    [7, 0.8],
    [13, 0.8],
  ] as const)('interpolates wrapped hue %s', (hue, expected) => {
    expect(hueToRgb(0.2, 0.8, hue)).toBeCloseTo(expected, 12);
  });

  it('normalizes non-finite interpolation inputs', () => {
    expect(hueToRgb(Number.NaN, 0.8, 0)).toBe(0);
    expect(hueToRgb(0.2, Number.POSITIVE_INFINITY, 2)).toBe(0);
    expect(hueToRgb(0.2, 0.8, Number.NaN)).toBeCloseTo(0.2, 12);
  });

  it.each([
    [0, 1, 0.5, { b: 0, g: 0, r: 255 }],
    [120, 1, 0.5, { b: 0, g: 255, r: 0 }],
    [240, 1, 0.5, { b: 255, g: 0, r: 0 }],
    [480, 1, 0.5, { b: 0, g: 255, r: 0 }],
    [-120, 1, 0.5, { b: 255, g: 0, r: 0 }],
    [0, 0, 0.25, { b: 63.75, g: 63.75, r: 63.75 }],
    [60, 0.4, 0.75, { b: 165.75, g: 216.75, r: 216.75 }],
  ] as const)(
    'converts HSL(%s, %s, %s)',
    (hue, saturation, lightness, expected) => {
      const actual = hslToRgb(hue, saturation, lightness);
      expect(actual.r).toBeCloseTo(expected.r, 12);
      expect(actual.g).toBeCloseTo(expected.g, 12);
      expect(actual.b).toBeCloseTo(expected.b, 12);
    },
  );

  it('clamps HSL components and replaces non-finite values', () => {
    expect(hslToRgb(0, 2, 0.5)).toEqual({ b: 0, g: 0, r: 255 });
    expect(hslToRgb(0, -1, 0.5)).toEqual({
      b: 127.5,
      g: 127.5,
      r: 127.5,
    });
    expect(hslToRgb(0, 1, 2)).toEqual({ b: 255, g: 255, r: 255 });
    expect(
      hslToRgb(Number.POSITIVE_INFINITY, Number.NaN, Number.NEGATIVE_INFINITY),
    ).toEqual({ b: 0, g: 0, r: 0 });
  });
});

describe('PPTX color modifiers', () => {
  const modifiers = [
    ['shade', applyShade, 0.5, '1a334c'],
    ['tint', applyTint, 0.5, '8cb2d9'],
    ['luminance offset', applyLumOff, 0.25, '79a6d2'],
    ['luminance multiplier', applyLumMod, 0.5, '1a334c'],
    ['hue multiplier', applyHueMod, 2, '999933'],
    ['saturation multiplier', applySatMod, 0.5, '4d667f'],
  ] as const;

  it.each(modifiers)(
    'applies %s with and without alpha',
    (_, apply, value, hex) => {
      expect(apply('336699', value)).toBe(hex);
      expect(apply('#33669980', value, true)).toBe(`${hex}80`);
    },
  );

  it.each(modifiers)(
    'treats a non-finite %s as a neutral operation',
    (_, apply) => {
      expect(apply('336699', Number.NaN)).toBe('336699');
      expect(apply('336699', Number.POSITIVE_INFINITY)).toBe('336699');
    },
  );

  it('clamps shade and tint multipliers', () => {
    expect(applyShade('336699', -1)).toBe('000000');
    expect(applyShade('336699', 2)).toBe('336699');
    expect(applyTint('336699', -1)).toBe('ffffff');
    expect(applyTint('336699', 2)).toBe('336699');
  });

  it('clamps luminance operations', () => {
    expect(applyLumOff('336699', -2)).toBe('000000');
    expect(applyLumOff('336699', 2)).toBe('ffffff');
    expect(applyLumMod('336699', -1)).toBe('000000');
    expect(applyLumMod('336699', 10)).toBe('ffffff');
  });

  it('wraps hue and clamps saturation operations', () => {
    expect(applyHueMod('336699', -1)).toBe('339966');
    expect(applyHueMod('336699', 4)).toBe('339933');
    expect(applySatMod('336699', -1)).toBe('666666');
    expect(applySatMod('336699', 10)).toBe('0066cc');
  });
});

describe('DrawingML preset colors', () => {
  it.each([
    ['red', 'ff0000'],
    [' AliceBlue ', 'f0f8ff'],
    ['burntSienna', 'ea7e5d'],
    ['dkBlue', '00008b'],
    ['ltGoldenrodYellow', 'fafad2'],
    ['medAquamarine', '66cdaa'],
    ['mediumBlue', '0000cd'],
  ] as const)('resolves %s case-insensitively', (name, expected) => {
    expect(getColorName2Hex(name)).toBe(expected);
    expect(getColorName2Hex(name.toUpperCase())).toBe(expected);
  });

  it.each(['', 'unknownColor', '#ffffff', 'rgb(1,2,3)'])(
    'rejects non-preset color %j',
    (name) => {
      expect(getColorName2Hex(name)).toBeUndefined();
    },
  );
});

import { describe, expect, it } from 'vitest';

import { angleToDegrees, numberToFixed, toHex } from '../../src/common/numbers';

describe('numeric normalization helpers', () => {
  it.each([
    [undefined, 0],
    [null, 0],
    ['', 0],
    [0, 0],
    ['0', 0],
    [60_000, 1],
    ['120000', 2],
    [-60_000, -1],
    [29_999, 0],
    [30_000, 1],
    [-30_000, -0],
  ])('converts OOXML angle %j to %s degrees', (value, expected) => {
    expect(angleToDegrees(value)).toBe(expected);
  });

  it.each(['not-a-number', Number.NaN, Number.POSITIVE_INFINITY])(
    'does not return a non-finite angle for %j',
    (value) => {
      expect(angleToDegrees(value)).toBe(0);
    },
  );

  it.each([
    [0, '00'],
    [1, '01'],
    [15, '0f'],
    [16, '10'],
    [127, '7f'],
    [255, 'ff'],
  ])('formats byte %s as %s', (value, expected) => {
    expect(toHex(value)).toBe(expected);
  });

  it('rounds to four fractional digits by default', () => {
    expect(numberToFixed(1.234_567)).toBe(1.2346);
    expect(numberToFixed(-1.234_567)).toBe(-1.2346);
  });

  it.each([
    [1.6, 0, 2],
    [1.25, 1, 1.3],
    [1.234_567, 2, 1.23],
    [1.234_567, 6, 1.234_567],
  ])('rounds %s to %s digits', (value, digits, expected) => {
    expect(numberToFixed(value, digits)).toBe(expected);
  });
});

import { describe, expect, it } from 'vitest';

import {
  degreesToAngle,
  pointsToEmu,
  pointsToFontSize,
} from '../../src/formats/pptx/writer/units';

describe('PowerPoint writer unit conversion', () => {
  it.each([
    [0, 0],
    [1, 12_700],
    [-1, -12_700],
    [72, 914_400],
    [0.5, 6_350],
    [0.00004, 1],
    [0.000039, 0],
  ])('converts %s points to %s EMUs', (points, expected) => {
    expect(pointsToEmu(points)).toBe(expected);
  });

  it.each([
    [0, 0],
    [1, 60_000],
    [-90, -5_400_000],
    [360, 21_600_000],
    [0.00001, 1],
    [0.000008, 0],
  ])('converts %s degrees to %s angle units', (degrees, expected) => {
    expect(degreesToAngle(degrees)).toBe(expected);
  });

  it.each([
    [0, 0],
    [1, 100],
    [12, 1_200],
    [10.006, 1_001],
    [10.004, 1_000],
  ])('converts %s point fonts to %s hundredths', (points, expected) => {
    expect(pointsToFontSize(points)).toBe(expected);
  });

  it.each([Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY])(
    'rejects non-finite point value %s',
    (value) => {
      expect(() => pointsToEmu(value)).toThrow(
        new RangeError('Point value must be finite'),
      );
    },
  );

  it('uses domain-specific error messages', () => {
    expect(() => degreesToAngle(Number.NaN)).toThrow(
      new RangeError('Degree value must be finite'),
    );
    expect(() => pointsToFontSize(Number.NaN)).toThrow(
      new RangeError('Font size must be finite'),
    );
  });

  it('rejects scaled values outside the safe OOXML integer range', () => {
    expect(() => pointsToEmu(1_000_000_000_000)).toThrow(
      new RangeError('Point value exceeds the safe OOXML integer range'),
    );
    expect(() => degreesToAngle(1_000_000_000_000)).toThrow(
      new RangeError('Degree value exceeds the safe OOXML integer range'),
    );
    expect(() => pointsToFontSize(1_000_000_000_000_000)).toThrow(
      new RangeError('Font size exceeds the safe OOXML integer range'),
    );
  });
});

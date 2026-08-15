import { describe, expect, it } from 'vitest';

import {
  EMUS_PER_INCH,
  EMUS_TO_POINTS,
  POINTS_PER_INCH,
  RATIO_EMUs_Points,
  RATIO_Inches_EMUs,
  RATIO_Inches_Points,
} from '../../src/common/ooxml/units';

describe('OOXML unit constants', () => {
  it('represents one inch consistently in EMUs and points', () => {
    expect(EMUS_PER_INCH).toBe(914_400);
    expect(POINTS_PER_INCH).toBe(72);
    expect(EMUS_PER_INCH * EMUS_TO_POINTS).toBe(POINTS_PER_INCH);
    expect(12_700 * EMUS_TO_POINTS).toBe(1);
  });

  it('keeps compatibility constants identical to their canonical values', () => {
    expect(RATIO_Inches_EMUs).toBe(EMUS_PER_INCH);
    expect(RATIO_Inches_Points).toBe(POINTS_PER_INCH);
    expect(RATIO_EMUs_Points).toBe(EMUS_TO_POINTS);
  });
});

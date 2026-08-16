import { EMUS_PER_INCH, POINTS_PER_INCH } from '../../../common/ooxml/units';

const ANGLE_UNITS_PER_DEGREE = 60_000;
const FONT_SIZE_UNITS_PER_POINT = 100;
const EMUS_PER_POINT = EMUS_PER_INCH / POINTS_PER_INCH;

function scaledInteger(
  value: number,
  multiplier: number,
  label: string,
): number {
  if (!Number.isFinite(value)) {
    throw new RangeError(`${label} must be finite`);
  }
  const result = Math.round(value * multiplier);
  if (!Number.isSafeInteger(result)) {
    throw new RangeError(`${label} exceeds the safe OOXML integer range`);
  }
  return result;
}

export function pointsToEmu(value: number): number {
  return scaledInteger(value, EMUS_PER_POINT, 'Point value');
}

export function degreesToAngle(value: number): number {
  return scaledInteger(value, ANGLE_UNITS_PER_DEGREE, 'Degree value');
}

export function pointsToFontSize(value: number): number {
  return scaledInteger(value, FONT_SIZE_UNITS_PER_POINT, 'Font size');
}

/** English Metric Units in one inch. */
export const EMUS_PER_INCH = 914_400;

/** Points in one inch. */
export const POINTS_PER_INCH = 72;

/** Convert an OOXML EMU value to points. */
export const EMUS_TO_POINTS = POINTS_PER_INCH / EMUS_PER_INCH;

/** @internal Compatibility names used by the upstream PPTX parser. */
export const RATIO_Inches_EMUs = EMUS_PER_INCH;
export const RATIO_Inches_Points = POINTS_PER_INCH;
export const RATIO_EMUs_Points = EMUS_TO_POINTS;

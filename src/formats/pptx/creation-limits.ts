export const MAX_POWERPOINT_CREATION_ELEMENTS = 5_000;
export const MAX_POWERPOINT_CREATION_MEDIA = 1_000;
export const MAX_POWERPOINT_CREATION_MEDIA_BYTES = 64 * 1024 * 1024;
export const MAX_POWERPOINT_CREATION_TOTAL_MEDIA_BYTES = 256 * 1024 * 1024;
export const MAX_POWERPOINT_CREATION_PARAGRAPHS = 10_000;
export const MAX_POWERPOINT_CREATION_SLIDES = 10_000;
export const MAX_POWERPOINT_CREATION_STRING_CODE_UNITS = 8 * 1024 * 1024;
export const MAX_POWERPOINT_CREATION_TEXT_NODES = 40_000;

export function isSupportedPowerPointCreationSlideCount(
  slideCount: number,
): boolean {
  return (
    Number.isSafeInteger(slideCount) &&
    slideCount >= 0 &&
    slideCount <= MAX_POWERPOINT_CREATION_SLIDES
  );
}

export const MAX_POWERPOINT_CREATION_SLIDES = 10_000;

export function isSupportedPowerPointCreationSlideCount(
  slideCount: number,
): boolean {
  return (
    Number.isSafeInteger(slideCount) &&
    slideCount >= 0 &&
    slideCount <= MAX_POWERPOINT_CREATION_SLIDES
  );
}

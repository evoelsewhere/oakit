const CONTENT_TYPES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const PRESENTATION_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml';
const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const SLIDE_LAYOUT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml';
const SLIDE_MASTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml';
const THEME_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.theme+xml';
const CHART_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

function assertSlideCount(slideCount: number): void {
  if (!Number.isSafeInteger(slideCount) || slideCount < 0) {
    throw new RangeError(
      'PowerPoint slide count must be a non-negative safe integer',
    );
  }
}

export function serializeContentTypes(
  slideCount: number,
  mediaTypes: readonly ('image/jpeg' | 'image/png')[] = [],
  chartCount = 0,
): string {
  assertSlideCount(slideCount);
  assertSlideCount(chartCount);
  const fixed =
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    (mediaTypes.includes('image/png')
      ? '<Default Extension="png" ContentType="image/png"/>'
      : '') +
    (mediaTypes.includes('image/jpeg')
      ? '<Default Extension="jpeg" ContentType="image/jpeg"/>'
      : '') +
    `<Override PartName="/ppt/presentation.xml" ContentType="${PRESENTATION_CONTENT_TYPE}"/>` +
    `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${SLIDE_MASTER_CONTENT_TYPE}"/>` +
    `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${SLIDE_LAYOUT_CONTENT_TYPE}"/>` +
    `<Override PartName="/ppt/theme/theme1.xml" ContentType="${THEME_CONTENT_TYPE}"/>`;
  const slides = Array.from(
    { length: slideCount },
    (_, index) =>
      `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="${SLIDE_CONTENT_TYPE}"/>`,
  ).join('');
  const charts = Array.from(
    { length: chartCount },
    (_, index) =>
      `<Override PartName="/ppt/charts/chart${index + 1}.xml" ContentType="${CHART_CONTENT_TYPE}"/>`,
  ).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CONTENT_TYPES_NAMESPACE}">${fixed}${slides}${charts}</Types>`;
}

import type { PptxSceneSize } from '../scene-types';
import { isSupportedPowerPointCreationSlideCount } from '../creation-limits';
import { pointsToEmu } from './units';

const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export function assertPowerPointSlideCount(slideCount: number): void {
  if (!isSupportedPowerPointCreationSlideCount(slideCount)) {
    throw new RangeError(
      'PowerPoint presentation slide count must be an integer from 0 through 10000',
    );
  }
}

function serializeSlideIds(slideCount: number): string {
  if (slideCount === 0) return '';
  const ids = Array.from(
    { length: slideCount },
    (_, index) => `<p:sldId id="${index + 256}" r:id="rId${index + 2}"/>`,
  ).join('');
  return `<p:sldIdLst>${ids}</p:sldIdLst>`;
}

export function serializePresentation(
  size: PptxSceneSize,
  slideCount: number,
): string {
  assertPowerPointSlideCount(slideCount);
  const slideIds = serializeSlideIds(slideCount);
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<p:presentation xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIPS_NAMESPACE}" xmlns:p="${PRESENTATION_NAMESPACE}">` +
    '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
    slideIds +
    `<p:sldSz cx="${pointsToEmu(size.width)}" cy="${pointsToEmu(size.height)}" type="custom"/>` +
    '<p:notesSz cx="6858000" cy="9144000"/>' +
    '<p:defaultTextStyle/>' +
    '</p:presentation>'
  );
}

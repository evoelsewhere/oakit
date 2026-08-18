import { assertPowerPointSlideCount } from './presentation';
import { serializeRelationships } from './relationships';

const OFFICE_RELATIONSHIP_BASE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

export function serializeRootRelationships(): string {
  return serializeRelationships([
    {
      id: 'rId1',
      target: 'ppt/presentation.xml',
      type: `${OFFICE_RELATIONSHIP_BASE}officeDocument`,
    },
  ]);
}

export function serializePresentationRelationships(slideCount: number): string {
  assertPowerPointSlideCount(slideCount);
  const slides = Array.from({ length: slideCount }, (_, index) => ({
    id: `rId${index + 2}`,
    target: `slides/slide${index + 1}.xml`,
    type: `${OFFICE_RELATIONSHIP_BASE}slide`,
  }));
  return serializeRelationships([
    {
      id: 'rId1',
      target: 'slideMasters/slideMaster1.xml',
      type: `${OFFICE_RELATIONSHIP_BASE}slideMaster`,
    },
    ...slides,
  ]);
}

export function serializeSlideMasterRelationships(): string {
  return serializeRelationships([
    {
      id: 'rId1',
      target: '../slideLayouts/slideLayout1.xml',
      type: `${OFFICE_RELATIONSHIP_BASE}slideLayout`,
    },
    {
      id: 'rId2',
      target: '../theme/theme1.xml',
      type: `${OFFICE_RELATIONSHIP_BASE}theme`,
    },
  ]);
}

export function serializeSlideLayoutRelationships(): string {
  return serializeRelationships([
    {
      id: 'rId1',
      target: '../slideMasters/slideMaster1.xml',
      type: `${OFFICE_RELATIONSHIP_BASE}slideMaster`,
    },
  ]);
}

export function serializeSlideRelationships(
  imageTargets: readonly string[] = [],
): string {
  const images = imageTargets.map((target, index) => ({
    id: `rId${index + 2}`,
    target,
    type: `${OFFICE_RELATIONSHIP_BASE}image`,
  }));
  return serializeRelationships([
    {
      id: 'rId1',
      target: '../slideLayouts/slideLayout1.xml',
      type: `${OFFICE_RELATIONSHIP_BASE}slideLayout`,
    },
    ...images,
  ]);
}

import type { PptxSceneDocument } from '../scene-types';
import { serializeContentTypes } from './content-types';
import {
  serializeMinimalSlideLayout,
  serializeMinimalSlideMaster,
} from './hierarchy';
import { createFieldIdAllocator } from './identifiers';
import {
  serializePresentationRelationships,
  serializeRootRelationships,
  serializeSlideLayoutRelationships,
  serializeSlideMasterRelationships,
  serializeSlideRelationships,
} from './package-relationships';
import {
  assertPowerPointSlideCount,
  serializePresentation,
} from './presentation';
import { serializeSlide } from './slide';
import { serializeMinimalTheme } from './theme';

export interface PptxSerializedPart {
  path: string;
  xml: string;
}

function fixedParts(scene: PptxSceneDocument): PptxSerializedPart[] {
  const slideCount = scene.slides.length;
  return [
    { path: '[Content_Types].xml', xml: serializeContentTypes(slideCount) },
    { path: '_rels/.rels', xml: serializeRootRelationships() },
    {
      path: 'ppt/presentation.xml',
      xml: serializePresentation(scene.size, slideCount),
    },
    {
      path: 'ppt/_rels/presentation.xml.rels',
      xml: serializePresentationRelationships(slideCount),
    },
    {
      path: 'ppt/slideMasters/slideMaster1.xml',
      xml: serializeMinimalSlideMaster(),
    },
    {
      path: 'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      xml: serializeSlideMasterRelationships(),
    },
    {
      path: 'ppt/slideLayouts/slideLayout1.xml',
      xml: serializeMinimalSlideLayout(),
    },
    {
      path: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
      xml: serializeSlideLayoutRelationships(),
    },
    { path: 'ppt/theme/theme1.xml', xml: serializeMinimalTheme() },
  ];
}

export function serializePowerPointParts(
  scene: PptxSceneDocument,
): PptxSerializedPart[] {
  assertPowerPointSlideCount(scene.slides.length);
  const parts = fixedParts(scene);
  const fieldIds = createFieldIdAllocator();
  scene.slides.forEach((slide, index) => {
    const slideNumber = index + 1;
    parts.push({
      path: `ppt/slides/slide${slideNumber}.xml`,
      xml: serializeSlide(slide, fieldIds),
    });
    parts.push({
      path: `ppt/slides/_rels/slide${slideNumber}.xml.rels`,
      xml: serializeSlideRelationships(),
    });
  });
  return parts;
}

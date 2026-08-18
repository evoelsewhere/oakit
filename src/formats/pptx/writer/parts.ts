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
  data: string | Uint8Array;
  path: string;
}

interface SerializedMedia {
  data: Uint8Array;
  key: string;
  mimeType: 'image/jpeg' | 'image/png';
  path: string;
}

function serializeMedia(scene: PptxSceneDocument): SerializedMedia[] {
  return scene.media.map((media, index) => ({
    data: new Uint8Array(media.data),
    key: media.key,
    mimeType: media.mimeType,
    path: `ppt/media/image${index + 1}.${media.mimeType === 'image/png' ? 'png' : 'jpeg'}`,
  }));
}

function fixedParts(
  scene: PptxSceneDocument,
  media: readonly SerializedMedia[],
): PptxSerializedPart[] {
  const slideCount = scene.slides.length;
  return [
    {
      data: serializeContentTypes(
        slideCount,
        media.map((item) => item.mimeType),
      ),
      path: '[Content_Types].xml',
    },
    { data: serializeRootRelationships(), path: '_rels/.rels' },
    {
      data: serializePresentation(scene.size, slideCount),
      path: 'ppt/presentation.xml',
    },
    {
      data: serializePresentationRelationships(slideCount),
      path: 'ppt/_rels/presentation.xml.rels',
    },
    {
      data: serializeMinimalSlideMaster(),
      path: 'ppt/slideMasters/slideMaster1.xml',
    },
    {
      data: serializeSlideMasterRelationships(),
      path: 'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    },
    {
      data: serializeMinimalSlideLayout(),
      path: 'ppt/slideLayouts/slideLayout1.xml',
    },
    {
      data: serializeSlideLayoutRelationships(),
      path: 'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    },
    { data: serializeMinimalTheme(), path: 'ppt/theme/theme1.xml' },
  ];
}

export function serializePowerPointParts(
  scene: PptxSceneDocument,
): PptxSerializedPart[] {
  assertPowerPointSlideCount(scene.slides.length);
  const media = serializeMedia(scene);
  const mediaByKey = new Map(media.map((item) => [item.key, item]));
  const parts = fixedParts(scene, media);
  parts.push(...media.map(({ data, path }) => ({ data, path })));
  const fieldIds = createFieldIdAllocator();
  scene.slides.forEach((slide, index) => {
    const slideNumber = index + 1;
    const imageElements = slide.elements.filter(
      (element) => element.type === 'image',
    );
    const imageTargets = imageElements.map((element) => {
      if (element.mediaKey === undefined) {
        throw new TypeError(
          `PowerPoint image element ${element.key} has no media key`,
        );
      }
      const item = mediaByKey.get(element.mediaKey);
      if (item === undefined) {
        throw new TypeError(
          `PowerPoint image element ${element.key} references missing media ${element.mediaKey}`,
        );
      }
      return `../media/${item.path.slice('ppt/media/'.length)}`;
    });
    const imageRelationships = new Map(
      imageElements.map((element, imageIndex) => [
        element.key,
        `rId${imageIndex + 2}`,
      ]),
    );
    parts.push({
      data: serializeSlide(slide, fieldIds, imageRelationships),
      path: `ppt/slides/slide${slideNumber}.xml`,
    });
    parts.push({
      data: serializeSlideRelationships(imageTargets),
      path: `ppt/slides/_rels/slide${slideNumber}.xml.rels`,
    });
  });
  return parts;
}

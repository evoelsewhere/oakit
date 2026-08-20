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
import { serializeChartPart } from './chart';
import type { PptxSceneChartElement } from '../scene-types';

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

interface SerializedChart {
  data: string;
  element: PptxSceneChartElement;
  path: string;
}

function flattenSceneElements(
  elements: readonly PptxSceneDocument['slides'][number]['elements'][number][],
): PptxSceneDocument['slides'][number]['elements'] {
  const result: PptxSceneDocument['slides'][number]['elements'] = [];
  for (const element of elements) {
    result.push(element);
    if (element.type === 'group') {
      result.push(...flattenSceneElements(element.elements));
    }
  }
  return result;
}

function serializeMedia(scene: PptxSceneDocument): SerializedMedia[] {
  return scene.media.map((media, index) => ({
    data: new Uint8Array(media.data),
    key: media.key,
    mimeType: media.mimeType,
    path: `ppt/media/image${index + 1}.${media.mimeType === 'image/png' ? 'png' : 'jpeg'}`,
  }));
}

function serializeCharts(scene: PptxSceneDocument): SerializedChart[] {
  const charts = scene.slides.flatMap((slide) =>
    flattenSceneElements(slide.elements).filter(
      (element): element is PptxSceneChartElement => element.type === 'chart',
    ),
  );
  return charts.map((element, index) => ({
    data: serializeChartPart(element, index + 1),
    element,
    path: `ppt/charts/chart${index + 1}.xml`,
  }));
}

function fixedParts(
  scene: PptxSceneDocument,
  media: readonly SerializedMedia[],
  charts: readonly SerializedChart[],
): PptxSerializedPart[] {
  const slideCount = scene.slides.length;
  return [
    {
      data: serializeContentTypes(
        slideCount,
        media.map((item) => item.mimeType),
        charts.length,
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
  const charts = serializeCharts(scene);
  const mediaByKey = new Map(media.map((item) => [item.key, item]));
  const chartByKey = new Map(charts.map((item) => [item.element.key, item]));
  const parts = fixedParts(scene, media, charts);
  parts.push(...media.map(({ data, path }) => ({ data, path })));
  parts.push(...charts.map(({ data, path }) => ({ data, path })));
  const fieldIds = createFieldIdAllocator();
  scene.slides.forEach((slide, index) => {
    const slideNumber = index + 1;
    const imageElements = flattenSceneElements(slide.elements).filter(
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
    const chartElements = flattenSceneElements(slide.elements).filter(
      (element): element is PptxSceneChartElement => element.type === 'chart',
    );
    const chartTargets = chartElements.map((element) => {
      const item = chartByKey.get(element.key) as SerializedChart;
      return `../charts/${item.path.slice('ppt/charts/'.length)}`;
    });
    const chartRelationships = new Map(
      chartElements.map((element, chartIndex) => [
        element.key,
        `rId${imageElements.length + chartIndex + 2}`,
      ]),
    );
    parts.push({
      data: serializeSlide(
        slide,
        fieldIds,
        imageRelationships,
        chartRelationships,
      ),
      path: `ppt/slides/slide${slideNumber}.xml`,
    });
    parts.push({
      data: serializeSlideRelationships(imageTargets, chartTargets),
      path: `ppt/slides/_rels/slide${slideNumber}.xml.rels`,
    });
  });
  return parts;
}

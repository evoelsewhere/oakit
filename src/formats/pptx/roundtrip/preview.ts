import type {
  PptxSceneDocument,
  PptxSceneSlide,
  PptxSceneTransform,
  PptxSceneUnsupportedElement,
} from '../scene-types';
import type { PptxDocument, PptxElement } from '../types';

function resolvedTransform(element: PptxElement): PptxSceneTransform {
  return {
    height: element.height,
    width: element.width,
    x: element.left,
    y: element.top,
    ...('isFlipH' in element ? { flipHorizontal: element.isFlipH } : {}),
    ...('isFlipV' in element ? { flipVertical: element.isFlipV } : {}),
    ...('rotate' in element ? { rotation: element.rotate } : {}),
  };
}

function previewText(element: PptxElement): string | undefined {
  return 'content' in element ? element.content : undefined;
}

function sceneElement(
  element: PptxElement,
  slideIndex: number,
  elementIndex: number,
): PptxSceneUnsupportedElement {
  const text = previewText(element);
  return {
    authored: {},
    feature: element.type,
    key: `slide-${slideIndex + 1}-element-${elementIndex + 1}`,
    ...(text === undefined ? {} : { previewText: text }),
    resolved: { hidden: false, transform: resolvedTransform(element) },
    type: 'unsupported',
  };
}

function sceneSlide(slide: PptxDocument['slides'][number], index: number) {
  const result: PptxSceneSlide = {
    elements: slide.elements.map((element, elementIndex) =>
      sceneElement(element, index, elementIndex),
    ),
    key: `slide-${index + 1}`,
  };
  return result;
}

export function createPowerPointRoundTripPreview(
  document: PptxDocument,
): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { ...document.size },
    slides: document.slides.map(sceneSlide),
    themes: [],
  };
}

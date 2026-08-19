import { decodeXmlEntities } from '../../../common/text/html';
import type {
  PptxSceneDocument,
  PptxSceneElement,
  PptxSceneImageElement,
  PptxSceneShapeElement,
  PptxSceneSlide,
  PptxSceneTextBodyProperties,
  PptxSceneTextElement,
  PptxSceneTransform,
  PptxSceneUnsupportedElement,
} from '../scene-types';
import type { Image, PptxDocument, PptxElement, Shape, Text } from '../types';
import { createPptxRoundTripTablePreview } from './table-preview';
import { createPptxRoundTripGroupPreview } from './group-preview';

function resolvedTransform(
  element: PptxElement,
): PptxSceneTransform | undefined {
  if (
    !Number.isFinite(element.left) ||
    !Number.isFinite(element.top) ||
    !Number.isFinite(element.width) ||
    element.width <= 0 ||
    !Number.isFinite(element.height) ||
    element.height <= 0
  ) {
    return undefined;
  }
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

export function plainTextFromPowerPointHtml(html: string): string {
  const withLineBreaks = html
    .replace(/<br\s*\/?\s*>/gi, '\n')
    .replace(/<\/(?:li|p)\s*>/gi, '\n');
  const withoutTags = withLineBreaks.replace(/<[^>]*>/g, '');
  return decodeXmlEntities(withoutTags.replace(/&nbsp;/gi, ' ')).replace(
    /\n+$/,
    '',
  );
}

function textBodyProperties(element: Text): PptxSceneTextBodyProperties {
  const anchor =
    element.vAlign === 'down'
      ? 'bottom'
      : element.vAlign === 'mid'
        ? 'center'
        : element.vAlign === 'dist'
          ? 'distributed'
          : element.vAlign === 'just'
            ? 'justified'
            : 'top';
  return {
    anchor,
    ...(element.autoFit === undefined ? {} : { autoFit: element.autoFit.type }),
    vertical: element.isVertical,
    wrap: element.wrap,
  };
}

function sceneTextElement(
  element: Text,
  slideIndex: number,
  elementIndex: number,
  keyOverride?: string,
): PptxSceneTextElement {
  const key =
    keyOverride ?? `slide-${slideIndex + 1}-element-${elementIndex + 1}`;
  const transform = resolvedTransform(element) as PptxSceneTransform;
  return {
    authored: {},
    key,
    name: element.name,
    resolved: {
      hidden: false,
      transform,
    },
    text: {
      body: textBodyProperties(element),
      paragraphs: [
        {
          children: [
            {
              key: `${key}-run-1`,
              text: plainTextFromPowerPointHtml(element.content),
              type: 'run',
            },
          ],
          key: `${key}-paragraph-1`,
        },
      ],
    },
    type: 'text',
  };
}

function sceneShapeElement(
  element: Shape,
  slideIndex: number,
  elementIndex: number,
  keyOverride?: string,
): PptxSceneShapeElement {
  const transform = resolvedTransform(element);
  return {
    authored: {},
    key: keyOverride ?? `slide-${slideIndex + 1}-element-${elementIndex + 1}`,
    name: element.name,
    resolved: {
      hidden: false,
      ...(transform === undefined ? {} : { transform }),
    },
    type: 'shape',
  };
}

function sceneImageElement(
  element: Image,
  slideIndex: number,
  elementIndex: number,
  keyOverride?: string,
): PptxSceneImageElement {
  const transform = resolvedTransform(element);
  return {
    authored: {},
    key: keyOverride ?? `slide-${slideIndex + 1}-element-${elementIndex + 1}`,
    resolved: {
      hidden: false,
      ...(transform === undefined ? {} : { transform }),
    },
    type: 'image',
  };
}

function sceneUnsupportedElement(
  element: PptxElement,
  slideIndex: number,
  elementIndex: number,
  keyOverride?: string,
): PptxSceneUnsupportedElement {
  const text = previewText(element);
  const transform = resolvedTransform(element);
  return {
    authored: {},
    feature: element.type,
    key: keyOverride ?? `slide-${slideIndex + 1}-element-${elementIndex + 1}`,
    ...(text === undefined ? {} : { previewText: text }),
    resolved: {
      hidden: false,
      ...(transform === undefined ? {} : { transform }),
    },
    type: 'unsupported',
  };
}

function sceneElement(
  element: PptxElement,
  slideIndex: number,
  elementIndex: number,
  keyOverride?: string,
): PptxSceneElement {
  if (element.type === 'text') {
    return sceneTextElement(element, slideIndex, elementIndex, keyOverride);
  }
  if (element.type === 'shape') {
    return plainTextFromPowerPointHtml(element.content) === ''
      ? sceneShapeElement(element, slideIndex, elementIndex, keyOverride)
      : sceneUnsupportedElement(element, slideIndex, elementIndex, keyOverride);
  }
  if (element.type === 'image') {
    return sceneImageElement(element, slideIndex, elementIndex, keyOverride);
  }
  if (element.type === 'table') {
    return (
      createPptxRoundTripTablePreview(
        element,
        slideIndex,
        elementIndex,
        plainTextFromPowerPointHtml,
        resolvedTransform,
        keyOverride,
      ) ??
      sceneUnsupportedElement(element, slideIndex, elementIndex, keyOverride)
    );
  }
  if (element.type === 'group') {
    return (
      createPptxRoundTripGroupPreview(
        element,
        slideIndex,
        elementIndex,
        {
          mapChild: (child, childIndex, key) =>
            sceneElement(child, slideIndex, childIndex, key),
          resolveTransform: resolvedTransform,
        },
        keyOverride,
      ) ??
      sceneUnsupportedElement(element, slideIndex, elementIndex, keyOverride)
    );
  }
  return sceneUnsupportedElement(
    element,
    slideIndex,
    elementIndex,
    keyOverride,
  );
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

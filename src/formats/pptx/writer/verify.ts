import { RATIO_EMUs_Points } from '../../../common/ooxml/units';
import { renderPptxDocumentToSvg } from '../render-svg';
import type { PptxSvgRenderResult } from '../render-types';
import type {
  PptxSceneDocument,
  PptxSceneTextElement,
  PptxSceneTextNode,
} from '../scene-types';
import { parse } from '../parser';
import type { PptxDocument, PptxParseOptions, Shape, Text } from '../types';
import { plainTextFromPowerPointHtml } from '../roundtrip/preview';
import { pointsToEmu } from './units';

type PptxCreationParser = (
  data: Uint8Array,
  options: PptxParseOptions,
) => Promise<PptxDocument>;

type PptxCreationRenderer = (document: PptxDocument) => PptxSvgRenderResult;

function expectedPointValue(value: number): number {
  return pointsToEmu(value) * RATIO_EMUs_Points;
}

function textNodeValue(node: PptxSceneTextNode): string {
  return node.type === 'break' ? '\n' : node.text;
}

function expectedPlainText(element: PptxSceneTextElement): string {
  return element.text.paragraphs
    .map((paragraph) => paragraph.children.map(textNodeValue).join(''))
    .join('\n');
}

function generatedTextElement(
  generated: PptxDocument['slides'][number]['elements'][number] | undefined,
  geometry: NonNullable<PptxSceneTextElement['authored']['geometry']>,
  location: string,
): Shape | Text {
  if (generated === undefined) {
    throw new Error(`Generated PowerPoint text element missing at ${location}`);
  }
  if (geometry === 'rect') {
    if (generated.type === 'text') return generated;
  } else if (generated.type === 'shape' && generated.shapType === geometry) {
    return generated;
  }
  throw new Error(`Generated PowerPoint text element missing at ${location}`);
}

function verifyTextElement(
  generated: PptxDocument['slides'][number]['elements'][number] | undefined,
  expected: PptxSceneTextElement,
  slideIndex: number,
  elementIndex: number,
): void {
  const location = `slide ${slideIndex + 1}, element ${elementIndex + 1}`;
  const geometry = expected.authored.geometry ?? 'rect';
  const textElement = generatedTextElement(generated, geometry, location);
  const transform = expected.authored.transform;
  if (transform === undefined) {
    throw new Error(
      `Expected PowerPoint authored transform missing at ${location}`,
    );
  }
  const generatedTransform = {
    flipHorizontal: textElement.isFlipH,
    flipVertical: textElement.isFlipV,
    height: textElement.height,
    rotation: textElement.rotate,
    width: textElement.width,
    x: textElement.left,
    y: textElement.top,
  };
  const expectedTransform = {
    flipHorizontal: transform.flipHorizontal ?? false,
    flipVertical: transform.flipVertical ?? false,
    height: expectedPointValue(transform.height),
    rotation: transform.rotation ?? 0,
    width: expectedPointValue(transform.width),
    x: expectedPointValue(transform.x),
    y: expectedPointValue(transform.y),
  };
  if (
    JSON.stringify(generatedTransform) !== JSON.stringify(expectedTransform)
  ) {
    throw new Error(`Generated PowerPoint transform mismatch at ${location}`);
  }
  const actualText = plainTextFromPowerPointHtml(textElement.content);
  if (actualText !== expectedPlainText(expected)) {
    throw new Error(`Generated PowerPoint text mismatch at ${location}`);
  }
}

function verifyRenderedSlides(
  document: PptxDocument,
  rendered: PptxSvgRenderResult,
): void {
  if (rendered.slides.length !== document.slides.length) {
    throw new Error(
      `Generated PowerPoint render count mismatch: expected ${document.slides.length}, received ${rendered.slides.length}`,
    );
  }
  for (const [index, slide] of rendered.slides.entries()) {
    if (
      slide.format !== 'svg' ||
      slide.mimeType !== 'image/svg+xml' ||
      slide.slideNumber !== index + 1 ||
      slide.width !== document.size.width ||
      slide.height !== document.size.height ||
      slide.data.byteLength === 0
    ) {
      throw new Error(
        `Generated PowerPoint visual invariant mismatch on slide ${index + 1}`,
      );
    }
    const source = new TextDecoder().decode(slide.data);
    if (
      !/^(?:<\?xml[^>]*\?>)?<svg\b/.test(source) ||
      /<(?:foreignObject|script)\b/i.test(source) ||
      /(?:href|src)=["'](?:blob|file|https?):/i.test(source)
    ) {
      throw new Error(
        `Generated PowerPoint unsafe SVG output on slide ${index + 1}`,
      );
    }
  }
}

export async function verifyPowerPointCreationWithParser(
  data: Uint8Array,
  scene: PptxSceneDocument,
  parseDocument: PptxCreationParser,
  renderDocument: PptxCreationRenderer = renderPptxDocumentToSvg,
): Promise<void> {
  const document = await parseDocument(data, {
    audioMode: 'none',
    errorMode: 'strict',
    imageMode: 'none',
    limits: {
      maxEntries: scene.slides.length * 2 + 9,
      maxSlides: Math.max(1, scene.slides.length),
    },
    videoMode: 'none',
  });
  if (document.slides.length !== scene.slides.length) {
    throw new Error(
      `Generated PowerPoint slide count mismatch: expected ${scene.slides.length}, received ${document.slides.length}`,
    );
  }
  const expectedWidth = expectedPointValue(scene.size.width);
  const expectedHeight = expectedPointValue(scene.size.height);
  if (
    document.size.width !== expectedWidth ||
    document.size.height !== expectedHeight
  ) {
    throw new Error(
      `Generated PowerPoint size mismatch: expected ${expectedWidth}x${expectedHeight}, received ${document.size.width}x${document.size.height}`,
    );
  }
  scene.slides.forEach((slide, index) => {
    const generated = document.slides[index];
    if (!generated || generated.elements.length !== slide.elements.length) {
      throw new Error(
        `Generated PowerPoint element count mismatch on slide ${index + 1}: expected ${slide.elements.length}, received ${generated?.elements.length ?? 0}`,
      );
    }
    slide.elements.forEach((element, elementIndex) => {
      if (element.type !== 'text') {
        throw new Error(
          `Expected PowerPoint text element missing at slide ${index + 1}, element ${elementIndex + 1}`,
        );
      }
      verifyTextElement(
        generated.elements[elementIndex],
        element,
        index,
        elementIndex,
      );
    });
  });
  verifyRenderedSlides(document, renderDocument(document));
}

export function verifyPowerPointCreation(
  data: Uint8Array,
  scene: PptxSceneDocument,
): Promise<void> {
  return verifyPowerPointCreationWithParser(data, scene, parse);
}

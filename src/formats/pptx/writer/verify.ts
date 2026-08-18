import { RATIO_EMUs_Points } from '../../../common/ooxml/units';
import { encodeBase64 } from '../../../common/binary/base64';
import { renderPptxDocumentToSvg } from '../render-svg';
import type { PptxSvgRenderResult } from '../render-types';
import type {
  PptxSceneDocument,
  PptxSceneImageElement,
  PptxSceneMedia,
  PptxSceneShapeElement,
  PptxSceneTableBorder,
  PptxSceneTableCell,
  PptxSceneTableElement,
  PptxSceneTextElement,
  PptxSceneTextNode,
} from '../scene-types';
import { parse } from '../parser';
import type {
  Image,
  PptxDocument,
  PptxParseOptions,
  Shape,
  Table,
  Text,
} from '../types';
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
  verifyElementTransform(textElement, expected, location);
  const actualText = plainTextFromPowerPointHtml(textElement.content);
  if (actualText !== expectedPlainText(expected)) {
    throw new Error(`Generated PowerPoint text mismatch at ${location}`);
  }
}

function verifyElementTransform(
  generated: Image | Shape | Table | Text,
  expected:
    | PptxSceneImageElement
    | PptxSceneShapeElement
    | PptxSceneTableElement
    | PptxSceneTextElement,
  location: string,
): void {
  const transform = expected.authored.transform;
  if (transform === undefined) {
    throw new Error(
      `Expected PowerPoint authored transform missing at ${location}`,
    );
  }
  const generatedTransform = {
    flipHorizontal: generated.isFlipH ?? false,
    flipVertical: generated.isFlipV ?? false,
    height: generated.height,
    rotation: generated.rotate ?? 0,
    width: generated.width,
    x: generated.left,
    y: generated.top,
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
}

function verifyTableBorder(
  generated: Table['data'][number][number]['borders'][keyof Table['data'][number][number]['borders']],
  expected: PptxSceneTableBorder | undefined,
  location: string,
): void {
  if (expected === undefined) {
    if (generated !== undefined) {
      throw new Error(
        `Generated PowerPoint table border mismatch at ${location}`,
      );
    }
    return;
  }
  if (
    generated?.borderColor !== expected.color ||
    generated.borderWidth !== expected.width ||
    generated.borderType !== (expected.style ?? 'solid')
  ) {
    throw new Error(
      `Generated PowerPoint table border mismatch at ${location}`,
    );
  }
}

function verifyTableCell(
  generated: Table['data'][number][number] | undefined,
  expected: PptxSceneTableCell,
  location: string,
): void {
  if (generated === undefined) {
    throw new Error(`Generated PowerPoint table cell missing at ${location}`);
  }
  const expectedText = expected.text.paragraphs
    .map((paragraph) => paragraph.children.map(textNodeValue).join(''))
    .join('\n');
  if (plainTextFromPowerPointHtml(generated.text) !== expectedText) {
    throw new Error(`Generated PowerPoint table text mismatch at ${location}`);
  }
  if (
    generated.fillColor !== expected.fillColor ||
    generated.colSpan !== expected.colSpan ||
    generated.rowSpan !== expected.rowSpan ||
    generated.hMerge !== (expected.hMerge ? 1 : undefined) ||
    generated.vMerge !== (expected.vMerge ? 1 : undefined)
  ) {
    throw new Error(`Generated PowerPoint table cell mismatch at ${location}`);
  }
  for (const key of ['bottom', 'left', 'right', 'top'] as const) {
    verifyTableBorder(
      generated.borders[key],
      expected.borders?.[key],
      `${location}, ${key} border`,
    );
  }
}

function verifyTableElement(
  generated: PptxDocument['slides'][number]['elements'][number] | undefined,
  expected: PptxSceneTableElement,
  slideIndex: number,
  elementIndex: number,
): void {
  const location = `slide ${slideIndex + 1}, element ${elementIndex + 1}`;
  if (generated?.type !== 'table') {
    throw new Error(`Generated PowerPoint table missing at ${location}`);
  }
  verifyElementTransform(generated, expected, location);
  const expectedColumns = expected.columns.map(expectedPointValue);
  const expectedRows = expected.rows.map((row) =>
    expectedPointValue(row.height),
  );
  if (
    JSON.stringify(generated.colWidths) !== JSON.stringify(expectedColumns) ||
    JSON.stringify(generated.rowHeights) !== JSON.stringify(expectedRows) ||
    generated.data.length !== expected.rows.length
  ) {
    throw new Error(`Generated PowerPoint table grid mismatch at ${location}`);
  }
  expected.rows.forEach((row, rowIndex) => {
    const generatedRow = generated.data[rowIndex];
    if (generatedRow?.length !== row.cells.length) {
      throw new Error(
        `Generated PowerPoint table row mismatch at ${location}, row ${rowIndex + 1}`,
      );
    }
    row.cells.forEach((cell, cellIndex) =>
      verifyTableCell(
        generatedRow[cellIndex],
        cell,
        `${location}, row ${rowIndex + 1}, cell ${cellIndex + 1}`,
      ),
    );
  });
}

function verifyImageElement(
  generated: PptxDocument['slides'][number]['elements'][number] | undefined,
  expected: PptxSceneImageElement,
  media: PptxSceneMedia | undefined,
  slideIndex: number,
  elementIndex: number,
): void {
  const location = `slide ${slideIndex + 1}, element ${elementIndex + 1}`;
  if (generated?.type !== 'image') {
    throw new Error(`Generated PowerPoint image missing at ${location}`);
  }
  if (media === undefined) {
    throw new Error(`Expected PowerPoint image media missing at ${location}`);
  }
  verifyElementTransform(generated, expected, location);
  const expectedBase64 = `data:${media.mimeType};base64,${encodeBase64(media.data)}`;
  if (generated.base64 !== expectedBase64) {
    throw new Error(`Generated PowerPoint image data mismatch at ${location}`);
  }
  if (generated.geom !== 'rect') {
    throw new Error(
      `Generated PowerPoint image geometry mismatch at ${location}`,
    );
  }
}

function verifyShapeElement(
  generated: PptxDocument['slides'][number]['elements'][number] | undefined,
  expected: PptxSceneShapeElement,
  slideIndex: number,
  elementIndex: number,
): void {
  const location = `slide ${slideIndex + 1}, element ${elementIndex + 1}`;
  const geometry = expected.authored.geometry ?? 'rect';
  if (generated?.type !== 'shape' || generated.shapType !== geometry) {
    throw new Error(`Generated PowerPoint shape missing at ${location}`);
  }
  verifyElementTransform(generated, expected, location);
  if (
    expected.authored.fillColor !== undefined &&
    (generated.fill?.type !== 'color' ||
      generated.fill.value !== expected.authored.fillColor)
  ) {
    throw new Error(`Generated PowerPoint shape fill mismatch at ${location}`);
  }
  if (
    expected.authored.lineColor !== undefined &&
    generated.borderColor !== expected.authored.lineColor
  ) {
    throw new Error(`Generated PowerPoint shape line mismatch at ${location}`);
  }
  if (
    expected.authored.lineWidth !== undefined &&
    generated.borderWidth !== expected.authored.lineWidth
  ) {
    throw new Error(
      `Generated PowerPoint shape line width mismatch at ${location}`,
    );
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
    imageMode: scene.media.length === 0 ? 'none' : 'base64',
    limits: {
      maxEntries: scene.slides.length * 2 + scene.media.length + 9,
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
      if (element.type === 'text') {
        verifyTextElement(
          generated.elements[elementIndex],
          element,
          index,
          elementIndex,
        );
      } else if (element.type === 'shape') {
        verifyShapeElement(
          generated.elements[elementIndex],
          element,
          index,
          elementIndex,
        );
      } else if (element.type === 'image') {
        verifyImageElement(
          generated.elements[elementIndex],
          element,
          scene.media.find((media) => media.key === element.mediaKey),
          index,
          elementIndex,
        );
      } else if (element.type === 'table') {
        verifyTableElement(
          generated.elements[elementIndex],
          element,
          index,
          elementIndex,
        );
      } else {
        throw new Error(
          `Expected PowerPoint text element missing at slide ${index + 1}, element ${elementIndex + 1}`,
        );
      }
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

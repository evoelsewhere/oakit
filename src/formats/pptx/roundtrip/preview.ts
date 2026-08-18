import { decodeXmlEntities } from '../../../common/text/html';
import type {
  PptxSceneDocument,
  PptxSceneElement,
  PptxSceneImageElement,
  PptxSceneShapeElement,
  PptxSceneSlide,
  PptxSceneTableBorder,
  PptxSceneTableCell,
  PptxSceneTableElement,
  PptxSceneTextBodyProperties,
  PptxSceneTextElement,
  PptxSceneTransform,
  PptxSceneUnsupportedElement,
} from '../scene-types';
import type {
  Border,
  Image,
  PptxDocument,
  PptxElement,
  Shape,
  Table,
  TableCell,
  Text,
} from '../types';

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
): PptxSceneTextElement {
  const key = `slide-${slideIndex + 1}-element-${elementIndex + 1}`;
  const transform = resolvedTransform(element);
  return {
    authored: {},
    key,
    name: element.name,
    resolved: {
      hidden: false,
      ...(transform === undefined ? {} : { transform }),
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
): PptxSceneShapeElement {
  const transform = resolvedTransform(element);
  return {
    authored: {},
    key: `slide-${slideIndex + 1}-element-${elementIndex + 1}`,
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
): PptxSceneImageElement {
  const transform = resolvedTransform(element);
  return {
    authored: {},
    key: `slide-${slideIndex + 1}-element-${elementIndex + 1}`,
    resolved: {
      hidden: false,
      ...(transform === undefined ? {} : { transform }),
    },
    type: 'image',
  };
}

function sceneTableBorder(border: Border): PptxSceneTableBorder {
  return {
    color: border.borderColor,
    style: border.borderType,
    width: border.borderWidth,
  };
}

function hasVisibleTableBorder(border: Border | undefined): border is Border {
  return border !== undefined && border.borderWidth > 0;
}

function sceneTableCell(
  cell: TableCell,
  tableKey: string,
  rowIndex: number,
  columnIndex: number,
): PptxSceneTableCell {
  const key = `${tableKey}-row-${rowIndex + 1}-cell-${columnIndex + 1}`;
  const properties = {
    ...(cell.fontBold === undefined ? {} : { bold: cell.fontBold }),
    ...(cell.fontColor === undefined ? {} : { color: cell.fontColor }),
  };
  return {
    borders: {
      ...(!hasVisibleTableBorder(cell.borders.bottom)
        ? {}
        : { bottom: sceneTableBorder(cell.borders.bottom) }),
      ...(!hasVisibleTableBorder(cell.borders.left)
        ? {}
        : { left: sceneTableBorder(cell.borders.left) }),
      ...(!hasVisibleTableBorder(cell.borders.right)
        ? {}
        : { right: sceneTableBorder(cell.borders.right) }),
      ...(!hasVisibleTableBorder(cell.borders.top)
        ? {}
        : { top: sceneTableBorder(cell.borders.top) }),
    },
    ...(cell.colSpan === undefined ? {} : { colSpan: cell.colSpan }),
    ...(cell.fillColor === undefined ? {} : { fillColor: cell.fillColor }),
    ...(cell.hMerge === undefined ? {} : { hMerge: cell.hMerge === 1 }),
    ...(cell.rowSpan === undefined ? {} : { rowSpan: cell.rowSpan }),
    text: {
      body: {
        anchor:
          cell.vAlign === 'down'
            ? 'bottom'
            : cell.vAlign === 'mid'
              ? 'center'
              : 'top',
      },
      paragraphs: [
        {
          children: [
            {
              key: `${key}-run-1`,
              ...(Object.keys(properties).length === 0 ? {} : { properties }),
              text: plainTextFromPowerPointHtml(cell.text),
              type: 'run',
            },
          ],
          key: `${key}-paragraph-1`,
        },
      ],
    },
    ...(cell.vMerge === undefined ? {} : { vMerge: cell.vMerge === 1 }),
  };
}

function isNativeTablePreview(element: Table): boolean {
  if (element.colWidths.length === 0 || element.data.length === 0) return false;
  if (
    element.colWidths.some((value) => !Number.isFinite(value) || value <= 0) ||
    element.rowHeights.length !== element.data.length ||
    element.rowHeights.some((value) => !Number.isFinite(value) || value <= 0) ||
    element.data.some((row) => row.length !== element.colWidths.length)
  ) {
    return false;
  }
  const transform = resolvedTransform(element);
  if (transform === undefined) return false;
  const columnWidth = element.colWidths.reduce(
    (total, value) => total + value,
    0,
  );
  const rowHeight = element.rowHeights.reduce(
    (total, value) => total + value,
    0,
  );
  return (
    Math.round(columnWidth * 12_700) === Math.round(transform.width * 12_700) &&
    Math.round(rowHeight * 12_700) === Math.round(transform.height * 12_700)
  );
}

function sceneTableElement(
  element: Table,
  slideIndex: number,
  elementIndex: number,
): PptxSceneTableElement {
  const key = `slide-${slideIndex + 1}-element-${elementIndex + 1}`;
  const transform = resolvedTransform(element);
  return {
    authored: {},
    columns: [...element.colWidths],
    key,
    ...(element.name === undefined ? {} : { name: element.name }),
    resolved: {
      hidden: false,
      ...(transform === undefined ? {} : { transform }),
    },
    rows: element.data.map((row, rowIndex) => ({
      cells: row.map((cell, columnIndex) =>
        sceneTableCell(cell, key, rowIndex, columnIndex),
      ),
      height: element.rowHeights[rowIndex] ?? 0,
    })),
    type: 'table',
  };
}

function sceneUnsupportedElement(
  element: PptxElement,
  slideIndex: number,
  elementIndex: number,
): PptxSceneUnsupportedElement {
  const text = previewText(element);
  const transform = resolvedTransform(element);
  return {
    authored: {},
    feature: element.type,
    key: `slide-${slideIndex + 1}-element-${elementIndex + 1}`,
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
): PptxSceneElement {
  if (element.type === 'text') {
    return sceneTextElement(element, slideIndex, elementIndex);
  }
  if (element.type === 'shape') {
    return plainTextFromPowerPointHtml(element.content) === ''
      ? sceneShapeElement(element, slideIndex, elementIndex)
      : sceneUnsupportedElement(element, slideIndex, elementIndex);
  }
  if (element.type === 'image') {
    return sceneImageElement(element, slideIndex, elementIndex);
  }
  if (element.type === 'table') {
    return isNativeTablePreview(element)
      ? sceneTableElement(element, slideIndex, elementIndex)
      : sceneUnsupportedElement(element, slideIndex, elementIndex);
  }
  return sceneUnsupportedElement(element, slideIndex, elementIndex);
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

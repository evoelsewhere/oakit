import { unsupportedPptxEdit } from './patch-error';
import { SaxesParser } from 'saxes';

const PRESENTATION_NAMESPACES = new Set([
  'http://purl.oclc.org/ooxml/presentationml/main',
  'http://schemas.openxmlformats.org/presentationml/2006/main',
]);
const DRAWING_NAMESPACES = new Set([
  'http://purl.oclc.org/ooxml/drawingml/main',
  'http://schemas.openxmlformats.org/drawingml/2006/main',
]);
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';

interface ShapeFrame {
  matchesShapeId: boolean;
  parentPresentationShape: ShapeFrame | undefined;
  presentationShape: boolean;
  start: number;
}

interface ShapeRange {
  end: number;
  start: number;
}

export interface PptxEditableShapeXml {
  drawingPrefix: string;
  markupPrefix: string | undefined;
  presentationPrefix: string;
  range: ShapeRange;
  shape: string;
}

function namespacePrefixes(xml: string): Map<string, string> {
  let rootLocalName: string | undefined;
  let rootNamespaces: Record<string, string> | undefined;
  const parser = new SaxesParser({ xmlns: true });
  parser.on('opentag', (tag) => {
    if (rootLocalName !== undefined) return;
    rootLocalName = tag.local;
    rootNamespaces = { ...tag.ns };
  });
  try {
    parser.write(xml).close();
  } catch (cause) {
    unsupportedPptxEdit(
      'PowerPoint text edit slide root is unsupported',
      cause,
    );
  }
  if (rootLocalName !== 'sld' || rootNamespaces === undefined) {
    unsupportedPptxEdit('PowerPoint text edit slide root is unsupported');
  }
  return new Map(
    Object.entries(rootNamespaces).map(([prefix, namespace]) => [
      namespace,
      prefix,
    ]),
  );
}

function requiredNamespacePrefix(
  prefixes: ReadonlyMap<string, string>,
  namespaces: ReadonlySet<string>,
  description: string,
): string {
  for (const namespace of namespaces) {
    const prefix = prefixes.get(namespace);
    if (prefix !== undefined) return prefix;
  }
  unsupportedPptxEdit(
    `PowerPoint text edit slide has no ${description} namespace`,
  );
}

function shapeRange(
  xml: string,
  shapeId: string,
  localName: 'graphicFrame' | 'pic' | 'sp',
  description: string,
): ShapeRange {
  const stack: ShapeFrame[] = [];
  const ranges: ShapeRange[] = [];
  let activePresentationShape: ShapeFrame | undefined;
  const parser = new SaxesParser({ xmlns: true });
  parser.on('opentag', (tag) => {
    const start = parser.position;
    const presentationShape =
      PRESENTATION_NAMESPACES.has(tag.uri) && tag.local === localName;
    if (PRESENTATION_NAMESPACES.has(tag.uri) && tag.local === 'cNvPr') {
      const value = (tag.attributes.id as { value?: string } | undefined)
        ?.value;
      if (value === shapeId && activePresentationShape !== undefined) {
        activePresentationShape.matchesShapeId = true;
      }
    }
    const frame: ShapeFrame = {
      matchesShapeId: false,
      parentPresentationShape: activePresentationShape,
      presentationShape,
      start,
    };
    stack.push(frame);
    if (presentationShape) activePresentationShape = frame;
  });
  parser.on('closetag', () => {
    const frame = stack.pop() as ShapeFrame;
    if (frame.presentationShape && frame.matchesShapeId) {
      ranges.push({ end: parser.position, start: frame.start });
    }
    activePresentationShape = frame.parentPresentationShape;
  });
  parser.write(xml).close();
  if (ranges.length !== 1) {
    unsupportedPptxEdit(
      `PowerPoint text edit requires one unique ${description} for id ${shapeId}`,
    );
  }
  return ranges[0] as ShapeRange;
}

export function qualifiedPptxName(prefix: string, localName: string): string {
  return prefix.length === 0 ? localName : `${prefix}:${localName}`;
}

export function escapePptxXmlPattern(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function pptxShapeHasElement(xml: string, qualified: string): boolean {
  return new RegExp(`<${escapePptxXmlPattern(qualified)}(?:\\s|>|/)`).test(xml);
}

function resolvePptxEditableElementXml(
  xml: string,
  shapeId: string,
  localName: 'graphicFrame' | 'pic' | 'sp',
  description: string,
): PptxEditableShapeXml {
  const prefixes = namespacePrefixes(xml);
  const presentationPrefix = requiredNamespacePrefix(
    prefixes,
    PRESENTATION_NAMESPACES,
    'PresentationML',
  );
  const drawingPrefix = requiredNamespacePrefix(
    prefixes,
    DRAWING_NAMESPACES,
    'DrawingML',
  );
  const range = shapeRange(xml, shapeId, localName, description);
  return {
    drawingPrefix,
    markupPrefix: prefixes.get(MARKUP_COMPATIBILITY_NAMESPACE),
    presentationPrefix,
    range,
    shape: xml.slice(range.start, range.end),
  };
}

export function resolvePptxEditableShapeXml(
  xml: string,
  shapeId: string,
): PptxEditableShapeXml {
  return resolvePptxEditableElementXml(xml, shapeId, 'sp', 'text shape');
}

export function resolvePptxEditablePictureXml(
  xml: string,
  shapeId: string,
): PptxEditableShapeXml {
  return resolvePptxEditableElementXml(xml, shapeId, 'pic', 'picture');
}

export function resolvePptxEditableGraphicFrameXml(
  xml: string,
  shapeId: string,
): PptxEditableShapeXml {
  return resolvePptxEditableElementXml(
    xml,
    shapeId,
    'graphicFrame',
    'graphic frame',
  );
}

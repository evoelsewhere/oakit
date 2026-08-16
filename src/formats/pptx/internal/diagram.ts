import type { XmlLookupValue } from '../../../common';
import type {
  PptxDiagramContent,
  PptxParserContext,
  PptxRelationshipMap,
} from './context';

import { getRelationshipPartUri, getTextByPathList } from '../../../common';
import { getTextNodeValue } from './text';

const STANDARD_RELATIONSHIP_PREFIX =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const STRICT_RELATIONSHIP_PREFIX =
  'http://purl.oclc.org/ooxml/officeDocument/relationships/';

function asArray(value: XmlLookupValue | undefined): XmlLookupValue[] {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function renameDrawingPrefixes(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(renameDrawingPrefixes);
  if (typeof value !== 'object' || value === null) return value;

  return Object.fromEntries(
    Object.entries(value).map(([key, child]) => [
      key.startsWith('dsp:') ? `p:${key.slice(4)}` : key,
      renameDrawingPrefixes(child),
    ]),
  );
}

function relationshipTypeName(type: string | undefined): string {
  if (!type) return '';
  if (type.startsWith(STANDARD_RELATIONSHIP_PREFIX)) {
    return type.slice(STANDARD_RELATIONSHIP_PREFIX.length);
  }
  if (type.startsWith(STRICT_RELATIONSHIP_PREFIX)) {
    return type.slice(STRICT_RELATIONSHIP_PREFIX.length);
  }
  return type;
}

async function getPartRelationships(
  context: PptxParserContext,
  ownerPart: string,
): Promise<PptxRelationshipMap> {
  const result: PptxRelationshipMap = {};
  const relationships = await context.xmlReader.read(
    getRelationshipPartUri(ownerPart),
  );
  const relationshipNodes = getTextByPathList<XmlLookupValue>(relationships, [
    'Relationships',
    'Relationship',
  ]);

  for (const relationship of asArray(relationshipNodes)) {
    const attributes =
      getTextByPathList<Record<string, string>>(relationship, ['attrs']) ?? {};
    const id = attributes.Id;
    const target = attributes.Target;
    if (!id || !target) continue;
    const resolvedTarget = context.xmlReader.resolveRelationshipTarget(
      ownerPart,
      target,
      attributes.TargetMode,
    );
    if (!resolvedTarget) continue;

    result[id] = {
      type: relationshipTypeName(attributes.Type),
      target: resolvedTarget,
    };
  }

  return result;
}

export async function loadDiagramFile(
  context: PptxParserContext,
  filename: string,
  transformDrawing = false,
): Promise<XmlLookupValue | null> {
  if (!filename) return null;

  const cacheKey = String(transformDrawing) + filename;
  if (Object.hasOwn(context.diagramFileCache, cacheKey)) {
    return context.diagramFileCache[cacheKey] ?? null;
  }

  let content: XmlLookupValue | null = await context.xmlReader.read(filename);
  if (content && transformDrawing) {
    content = renameDrawingPrefixes(content) as XmlLookupValue;
  }

  context.diagramFileCache[cacheKey] = content;
  return content;
}

export function getDiagramDrawingRelId(
  dataContent: XmlLookupValue | null,
): string {
  const extensionNodes = getTextByPathList<XmlLookupValue>(dataContent, [
    'dgm:dataModel',
    'dgm:extLst',
    'a:ext',
  ]);
  for (const extension of asArray(extensionNodes)) {
    const relationshipId = getTextByPathList<string>(extension, [
      'dsp:dataModelExt',
      'attrs',
      'relId',
    ]);
    if (relationshipId) return relationshipId;
  }
  return '';
}

function relationshipTarget(
  relationships: PptxRelationshipMap,
  relationshipId: string | undefined,
): string | undefined {
  return relationshipId ? relationships[relationshipId]?.target : undefined;
}

export async function getDiagramNodeContext(
  node: XmlLookupValue,
  context: PptxParserContext,
  relationships: PptxRelationshipMap = context.slideResObj,
): Promise<PptxParserContext> {
  const relationshipIds =
    getTextByPathList<Record<string, string>>(node, [
      'a:graphic',
      'a:graphicData',
      'dgm:relIds',
      'attrs',
    ]) ?? {};
  const diagramContent: PptxDiagramContent = {
    data: null,
    layout: null,
    quickStyle: null,
    colors: null,
    drawing: null,
  };
  const diagramResObj: PptxRelationshipMap = {};

  const dataTarget = relationshipTarget(relationships, relationshipIds['r:dm']);
  const layoutTarget = relationshipTarget(
    relationships,
    relationshipIds['r:lo'],
  );
  const quickStyleTarget = relationshipTarget(
    relationships,
    relationshipIds['r:qs'],
  );
  const colorsTarget = relationshipTarget(
    relationships,
    relationshipIds['r:cs'],
  );

  if (dataTarget)
    diagramContent.data = await loadDiagramFile(context, dataTarget);
  if (layoutTarget)
    diagramContent.layout = await loadDiagramFile(context, layoutTarget);
  if (quickStyleTarget)
    diagramContent.quickStyle = await loadDiagramFile(
      context,
      quickStyleTarget,
    );
  if (colorsTarget)
    diagramContent.colors = await loadDiagramFile(context, colorsTarget);

  const drawingRelationshipId = getDiagramDrawingRelId(diagramContent.data);
  const dataRelationships = dataTarget
    ? await getPartRelationships(context, dataTarget)
    : {};
  const drawingTarget =
    dataRelationships[drawingRelationshipId]?.target ??
    relationshipTarget(relationships, drawingRelationshipId);
  let drawing: XmlLookupValue | null = null;

  if (drawingTarget) {
    drawing = await loadDiagramFile(context, drawingTarget, true);
    diagramContent.drawing = drawing;
    Object.assign(
      diagramResObj,
      await getPartRelationships(context, drawingTarget),
    );
  }

  return {
    ...context,
    digramFileContent: drawing ?? undefined,
    diagramResObj,
    diagramContent,
  } as PptxParserContext;
}

export function getSmartArtTextData(dataContent: XmlLookupValue): string[] {
  const result: string[] = [];
  const points = getTextByPathList<XmlLookupValue>(dataContent, [
    'dgm:dataModel',
    'dgm:ptLst',
    'dgm:pt',
  ]);

  for (const point of asArray(points)) {
    const textBody = getTextByPathList<XmlLookupValue>(point, ['dgm:t']);

    let nodeText = '';
    const paragraphs = getTextByPathList<XmlLookupValue>(textBody, ['a:p']);
    for (const paragraph of asArray(paragraphs)) {
      const runs = getTextByPathList<XmlLookupValue>(paragraph, ['a:r']);
      for (const run of asArray(runs)) {
        const textNode = getTextByPathList<XmlLookupValue>(run, ['a:t']);
        const text = getTextNodeValue(textNode);
        if (text) nodeText += text;
      }
      nodeText += '\n';
    }

    const cleanText = nodeText.trim();
    if (cleanText) result.push(cleanText);
  }

  return result;
}

import type { XmlLookupValue } from '../../../common';
import type {
  PptxDiagramContent,
  PptxParserContext,
  PptxRelationshipMap,
} from './context';

import {
  getRelationshipPartUri,
  getTextByPathList,
  resolveRelationshipTarget,
} from '../../../common';
import { getTextNodeValue } from './text';

function asArray(value: XmlLookupValue | undefined): XmlLookupValue[] {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

export async function loadDiagramFile(
  context: PptxParserContext,
  filename: string,
  transformDrawing = false,
): Promise<XmlLookupValue | null> {
  if (!filename) return null;

  const cacheKey = `${transformDrawing ? 'drawing:' : 'xml:'}${filename}`;
  const cached = context.diagramFileCache[cacheKey];
  if (cached) return cached;

  let content: XmlLookupValue | null = await context.xmlReader.read(filename);
  if (content && transformDrawing) {
    content = JSON.parse(
      JSON.stringify(content).replace(/dsp:/g, 'p:'),
    ) as XmlLookupValue;
  }

  context.diagramFileCache[cacheKey] = content;
  return content;
}

export function getDiagramDrawingRelId(dataContent: XmlLookupValue): string {
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
  context: PptxParserContext,
  relationshipId: string | undefined,
): string | undefined {
  return relationshipId
    ? context.slideResObj[relationshipId]?.target
    : undefined;
}

export async function getDiagramNodeContext(
  node: XmlLookupValue,
  context: PptxParserContext,
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

  const dataTarget = relationshipTarget(context, relationshipIds['r:dm']);
  const layoutTarget = relationshipTarget(context, relationshipIds['r:lo']);
  const quickStyleTarget = relationshipTarget(context, relationshipIds['r:qs']);
  const colorsTarget = relationshipTarget(context, relationshipIds['r:cs']);

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

  const drawingRelationshipId = diagramContent.data
    ? getDiagramDrawingRelId(diagramContent.data)
    : '';
  const drawingTarget = relationshipTarget(context, drawingRelationshipId);
  let drawing: XmlLookupValue | null = null;

  if (drawingTarget) {
    drawing = await loadDiagramFile(context, drawingTarget, true);
    diagramContent.drawing = drawing;

    const drawingName = drawingTarget.split('/').pop();
    if (drawingName) {
      const relationshipsFilename = getRelationshipPartUri(drawingTarget);
      const relationships = await context.xmlReader.read(relationshipsFilename);
      const relationshipNodes = getTextByPathList<XmlLookupValue>(
        relationships,
        ['Relationships', 'Relationship'],
      );

      for (const relationship of asArray(relationshipNodes)) {
        const attributes =
          getTextByPathList<Record<string, string>>(relationship, ['attrs']) ??
          {};
        const id = attributes.Id;
        const target = attributes.Target;
        if (!id || !target) continue;

        diagramResObj[id] = {
          type: (attributes.Type ?? '').replace(
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/',
            '',
          ),
          target: resolveRelationshipTarget(
            drawingTarget,
            target,
            attributes.TargetMode,
          ),
        };
      }
    }
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
    if (!textBody) continue;

    let nodeText = '';
    const paragraphs = getTextByPathList<XmlLookupValue>(textBody, ['a:p']);
    for (const paragraph of asArray(paragraphs)) {
      const runs = getTextByPathList<XmlLookupValue>(paragraph, ['a:r']);
      for (const run of asArray(runs)) {
        const textNode = getTextByPathList<XmlLookupValue>(run, ['a:t']);
        if (!textNode) continue;
        const text = getTextNodeValue(textNode);
        if (text) nodeText += text;
      }
      if (nodeText) nodeText += '\n';
    }

    const cleanText = nodeText.trim();
    if (cleanText) result.push(cleanText);
  }

  return result;
}

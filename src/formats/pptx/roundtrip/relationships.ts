import { decodeXmlEntities } from '../../../common/text/html';
import { resolveRelationshipTarget } from '../../../common/opc/part-uri';
import { readXmlFileResult } from '../../../common/xml/read-xml';
import type { ResolvedPptxResourceLimits } from '../internal/resource-limits';
import { unsupportedPptxEdit } from './patch-error';
import JSZip from 'jszip';

const PRESENTATION_PART = 'ppt/presentation.xml';
const PRESENTATION_RELATIONSHIPS_PART = 'ppt/_rels/presentation.xml.rels';

function isRecord(value: unknown): value is Record<string, unknown> {
  return value instanceof Object;
}

function valueAt(value: unknown, path: readonly string[]): unknown {
  let current = value;
  for (const key of path) {
    if (!isRecord(current)) return undefined;
    current = current[key];
  }
  return current;
}

function asArray(value: unknown): readonly unknown[] {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}

function attributes(value: unknown): Record<string, string> {
  const candidate = valueAt(value, ['attrs']);
  if (!isRecord(candidate)) return {};
  const result: Record<string, string> = {};
  for (const [key, item] of Object.entries(candidate)) {
    if (typeof item === 'string') result[key] = item;
  }
  return result;
}

async function requiredXmlTree(
  archive: JSZip,
  part: string,
  limits: ResolvedPptxResourceLimits,
): Promise<unknown> {
  const result = await readXmlFileResult(archive, part, {
    maxBytes: limits.maxXmlBytes,
    maxDepth: limits.maxXmlDepth,
    maxNodes: limits.maxXmlNodes,
  });
  if (result.status !== 'ok') {
    unsupportedPptxEdit(
      `PowerPoint text edit cannot read required part ${part}`,
      result.status === 'error' ? result.error : undefined,
    );
  }
  return result.value;
}

export async function resolvePptxSlideParts(
  archive: JSZip,
  limits: ResolvedPptxResourceLimits,
): Promise<string[]> {
  const [presentation, relationships] = await Promise.all([
    requiredXmlTree(archive, PRESENTATION_PART, limits),
    requiredXmlTree(archive, PRESENTATION_RELATIONSHIPS_PART, limits),
  ]);
  const targets = new Map<string, string>();
  for (const relationship of asArray(
    valueAt(relationships, ['Relationships', 'Relationship']),
  )) {
    const values = attributes(relationship);
    if (!values.Id || !values.Target || !values.Type?.endsWith('/slide')) {
      continue;
    }
    let target: string;
    try {
      target = resolveRelationshipTarget(
        PRESENTATION_PART,
        decodeXmlEntities(values.Target),
        values.TargetMode,
      );
    } catch (cause) {
      unsupportedPptxEdit(
        'PowerPoint text edit encountered an unsafe slide relationship',
        cause,
      );
    }
    if (archive.file(target) === null) continue;
    targets.set(values.Id, target);
  }

  const result: string[] = [];
  for (const slideId of asArray(
    valueAt(presentation, ['p:presentation', 'p:sldIdLst', 'p:sldId']),
  )) {
    const relationshipId = attributes(slideId)['r:id'];
    const target = relationshipId ? targets.get(relationshipId) : undefined;
    if (target === undefined) {
      unsupportedPptxEdit(
        'PowerPoint text edit cannot resolve the presentation slide order',
      );
    }
    result.push(target);
  }
  return result;
}

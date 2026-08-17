import { readZipEntryBytes } from '../../../common/archive/read-entry';
import { decodeXmlEntities } from '../../../common/text/html';
import { resolveRelationshipTarget } from '../../../common/opc/part-uri';
import { readXmlFileResult } from '../../../common/xml/read-xml';
import {
  assertXmlComplexity,
  decodeXmlBytes,
} from '../../../common/xml/validate';
import {
  assertPptxArchiveWithinLimits,
  assertPptxInputWithinLimits,
  type ResolvedPptxResourceLimits,
} from '../internal/resource-limits';
import type { PptxDocument } from '../types';
import { PptxWriteError } from '../write-error';
import { escapeXmlText } from '../writer/xml';
import { degreesToAngle, pointsToEmu } from '../writer/units';
import type {
  PptxRoundTripOperation,
  PptxRoundTripReplaceTextOperation,
  PptxRoundTripSetTransformOperation,
} from './types';
import JSZip from 'jszip';

const PRESENTATION_PART = 'ppt/presentation.xml';
const PRESENTATION_RELATIONSHIPS_PART = 'ppt/_rels/presentation.xml.rels';
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
const TARGET_KEY_PATTERN = /^slide-([1-9]\d*)-element-([1-9]\d*)-run-1$/;
const TRANSFORM_TARGET_KEY_PATTERN = /^slide-([1-9]\d*)-element-([1-9]\d*)$/;
const OFFICE_ESCAPE_PATTERN = /_x[0-9a-f]{4}_/i;
const XML_TAG_PATTERN =
  /<(\/?)(([A-Za-z_][\w.-]*):)?([A-Za-z_][\w.-]*)((?:\s+[A-Za-z_][\w.:-]*(?:\s*=\s*(?:"[^"]*"|'[^']*'))?)*\s*)(\/?)>/g;

interface ShapeFrame {
  matchesShapeId: boolean;
  name: string;
  start: number;
}

interface ShapeRange {
  end: number;
  start: number;
}

interface TextTarget {
  elementIndex: number;
  shapeId: string;
  slideIndex: number;
}

function activeShapeFrame(
  stack: ShapeFrame[],
  shapeName: string,
): ShapeFrame | undefined {
  for (let index = stack.length - 1; index >= 0; index -= 1) {
    const frame = stack[index];
    if (frame?.name === shapeName) return frame;
  }
}

export interface PptxPatchedPackage {
  copiedPartCount: number;
  data: Uint8Array;
  patchedPartCount: number;
}

function unsupportedEdit(message: string, cause?: unknown): never {
  throw new PptxWriteError('unsupported-edit-operation', message, { cause });
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
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
  if (value === undefined || value === null) return [];
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
    unsupportedEdit(
      `PowerPoint text edit cannot read required part ${part}`,
      result.status === 'error' ? result.error : undefined,
    );
  }
  return result.value;
}

async function resolveSlideParts(
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
    if (
      !values.Id ||
      !values.Target ||
      !values.Type?.endsWith('/slide') ||
      values.TargetMode?.toLowerCase() === 'external'
    ) {
      continue;
    }
    let target: string | null;
    try {
      target = resolveRelationshipTarget(
        PRESENTATION_PART,
        decodeXmlEntities(values.Target),
        values.TargetMode,
      );
    } catch (cause) {
      unsupportedEdit(
        'PowerPoint text edit encountered an unsafe slide relationship',
        cause,
      );
    }
    if (target !== null && archive.file(target) !== null) {
      targets.set(values.Id, target);
    }
  }

  const result: string[] = [];
  for (const slideId of asArray(
    valueAt(presentation, ['p:presentation', 'p:sldIdLst', 'p:sldId']),
  )) {
    const relationshipId = attributes(slideId)['r:id'];
    const target = relationshipId ? targets.get(relationshipId) : undefined;
    if (target === undefined) {
      unsupportedEdit(
        'PowerPoint text edit cannot resolve the presentation slide order',
      );
    }
    result.push(target);
  }
  return result;
}

function textTarget(
  operation: PptxRoundTripReplaceTextOperation,
  document: PptxDocument,
): TextTarget {
  const match = TARGET_KEY_PATTERN.exec(operation.targetKey);
  if (match === null) {
    unsupportedEdit(
      'PowerPoint text edit target is not a supported slide text run key',
    );
  }
  const slideIndex = Number(match[1]) - 1;
  const elementIndex = Number(match[2]) - 1;
  if (
    !Number.isSafeInteger(slideIndex) ||
    !Number.isSafeInteger(elementIndex)
  ) {
    unsupportedEdit('PowerPoint text edit target index is unsafe');
  }
  const element = document.slides[slideIndex]?.elements[elementIndex];
  if (element?.type !== 'text' || typeof element.id !== 'string') {
    unsupportedEdit(
      'PowerPoint text edit target is not a slide-owned text element',
    );
  }
  return { elementIndex, shapeId: element.id, slideIndex };
}

function transformTarget(
  operation: PptxRoundTripSetTransformOperation,
  document: PptxDocument,
): TextTarget {
  const match = TRANSFORM_TARGET_KEY_PATTERN.exec(operation.targetKey);
  if (match === null) {
    unsupportedEdit(
      'PowerPoint transform target is not a supported slide text element key',
    );
  }
  const slideIndex = Number(match[1]) - 1;
  const elementIndex = Number(match[2]) - 1;
  if (
    !Number.isSafeInteger(slideIndex) ||
    !Number.isSafeInteger(elementIndex)
  ) {
    unsupportedEdit('PowerPoint transform target index is unsafe');
  }
  const element = document.slides[slideIndex]?.elements[elementIndex];
  if (element?.type !== 'text' || typeof element.id !== 'string') {
    unsupportedEdit(
      'PowerPoint transform target is not a slide-owned text element',
    );
  }
  return { elementIndex, shapeId: element.id, slideIndex };
}

function namespacePrefixes(xml: string): Map<string, string> {
  const root =
    /<(([A-Za-z_][\w.-]*):)?sld\b((?:\s+[A-Za-z_][\w.:-]*\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*>/.exec(
      xml,
    );
  if (root === null) {
    unsupportedEdit('PowerPoint text edit slide root is unsupported');
  }
  const prefixes = new Map<string, string>();
  const attributesText = root[3] ?? '';
  const declarations =
    /(?:^|\s)xmlns(?::([A-Za-z_][\w.-]*))?\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
  for (const match of attributesText.matchAll(declarations)) {
    prefixes.set(match[2] ?? match[3] ?? '', match[1] ?? '');
  }
  return prefixes;
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
  unsupportedEdit(`PowerPoint text edit slide has no ${description} namespace`);
}

function qualifiedName(prefix: string, localName: string): string {
  return prefix.length === 0 ? localName : `${prefix}:${localName}`;
}

function attributeValue(attributesText: string, name: string): string | null {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const pattern = new RegExp(
    `(?:^|\\s)${escaped}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`,
  );
  const match = pattern.exec(attributesText);
  return match === null ? null : decodeXmlEntities(match[1] ?? match[2] ?? '');
}

function shapeRange(
  xml: string,
  presentationPrefix: string,
  shapeId: string,
): ShapeRange {
  const shapeName = qualifiedName(presentationPrefix, 'sp');
  const propertiesName = qualifiedName(presentationPrefix, 'cNvPr');
  const stack: ShapeFrame[] = [];
  const ranges: ShapeRange[] = [];
  XML_TAG_PATTERN.lastIndex = 0;
  for (const match of xml.matchAll(XML_TAG_PATTERN)) {
    const closing = match[1] === '/';
    const name = `${match[3] === undefined ? '' : `${match[3]}:`}${match[4] ?? ''}`;
    const attributesText = match[5] ?? '';
    const selfClosing = match[6] === '/';
    if (closing) {
      const frame = stack.pop();
      if (
        frame?.name === shapeName &&
        name === shapeName &&
        frame.matchesShapeId
      ) {
        ranges.push({
          end: (match.index ?? 0) + match[0].length,
          start: frame.start,
        });
      }
      continue;
    }
    if (
      name === propertiesName &&
      attributeValue(attributesText, 'id') === shapeId
    ) {
      const owner = activeShapeFrame(stack, shapeName);
      if (owner !== undefined) owner.matchesShapeId = true;
    }
    if (!selfClosing) {
      stack.push({
        matchesShapeId: false,
        name,
        start: match.index ?? 0,
      });
    }
  }
  if (ranges.length !== 1) {
    unsupportedEdit(
      `PowerPoint text edit requires one unique text shape for id ${shapeId}`,
    );
  }
  const range = ranges[0];
  if (range === undefined) {
    unsupportedEdit('PowerPoint text edit shape range disappeared');
  }
  return range;
}

function escapedPattern(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function hasElement(xml: string, qualified: string): boolean {
  return new RegExp(`<${escapedPattern(qualified)}(?:\\s|>|/)`).test(xml);
}

function patchShapeText(
  xml: string,
  shapeId: string,
  operation: PptxRoundTripReplaceTextOperation,
): string {
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
  const range = shapeRange(xml, presentationPrefix, shapeId);
  const shape = xml.slice(range.start, range.end);
  const markupPrefix = prefixes.get(MARKUP_COMPATIBILITY_NAMESPACE);
  if (
    (markupPrefix !== undefined &&
      hasElement(shape, qualifiedName(markupPrefix, 'AlternateContent'))) ||
    hasElement(shape, qualifiedName(presentationPrefix, 'extLst')) ||
    hasElement(shape, qualifiedName(drawingPrefix, 'extLst'))
  ) {
    unsupportedEdit(
      'PowerPoint text edit target contains unsupported compatibility markup',
    );
  }
  if (
    hasElement(shape, qualifiedName(drawingPrefix, 'br')) ||
    hasElement(shape, qualifiedName(drawingPrefix, 'fld'))
  ) {
    unsupportedEdit(
      'PowerPoint text edit target must contain one plain text run',
    );
  }

  const textName = qualifiedName(drawingPrefix, 't');
  const escapedTextName = escapedPattern(textName);
  const textPattern = new RegExp(
    `<${escapedTextName}((?:\\s+[A-Za-z_][\\w.:-]*\\s*=\\s*(?:"[^"]*"|'[^']*'))*)\\s*>([^<]*)<\\/${escapedTextName}\\s*>`,
    'g',
  );
  const matches = [...shape.matchAll(textPattern)];
  if (matches.length !== 1) {
    unsupportedEdit(
      'PowerPoint text edit target must contain exactly one text node',
    );
  }
  const match = matches[0];
  if (match === undefined) {
    unsupportedEdit('PowerPoint text edit text node disappeared');
  }
  const sourceText = decodeXmlEntities(match[2] ?? '');
  if (
    OFFICE_ESCAPE_PATTERN.test(sourceText) ||
    sourceText !== operation.expectedText
  ) {
    unsupportedEdit(
      'PowerPoint text edit source XML does not match its preview precondition',
    );
  }
  const originalAttributes = match[1] ?? '';
  const attributesWithoutSpace = originalAttributes.replace(
    /\s+xml:space\s*=\s*(?:"[^"]*"|'[^']*')/gi,
    '',
  );
  const replacement = `<${textName}${attributesWithoutSpace} xml:space="preserve">${escapeXmlText(operation.value)}</${textName}>`;
  const matchStart = range.start + (match.index ?? 0);
  const matchEnd = matchStart + match[0].length;
  return `${xml.slice(0, matchStart)}${replacement}${xml.slice(matchEnd)}`;
}

function integerAttribute(attributesText: string, name: string): number {
  const value = attributeValue(attributesText, name);
  if (value === null || !/^-?\d+$/.test(value)) {
    unsupportedEdit(`PowerPoint transform ${name} attribute is invalid`);
  }
  const result = Number(value);
  if (!Number.isSafeInteger(result)) {
    unsupportedEdit(`PowerPoint transform ${name} attribute is unsafe`);
  }
  return result;
}

function optionalIntegerAttribute(
  attributesText: string,
  name: string,
): number {
  return attributeValue(attributesText, name) === null
    ? 0
    : integerAttribute(attributesText, name);
}

function booleanAttribute(attributesText: string, name: string): boolean {
  const value = attributeValue(attributesText, name);
  if (value === null || value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  unsupportedEdit(`PowerPoint transform ${name} attribute is invalid`);
}

function patchShapeTransform(
  xml: string,
  shapeId: string,
  operation: PptxRoundTripSetTransformOperation,
): string {
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
  const range = shapeRange(xml, presentationPrefix, shapeId);
  const shape = xml.slice(range.start, range.end);
  const markupPrefix = prefixes.get(MARKUP_COMPATIBILITY_NAMESPACE);
  if (
    (markupPrefix !== undefined &&
      hasElement(shape, qualifiedName(markupPrefix, 'AlternateContent'))) ||
    hasElement(shape, qualifiedName(presentationPrefix, 'extLst')) ||
    hasElement(shape, qualifiedName(drawingPrefix, 'extLst'))
  ) {
    unsupportedEdit(
      'PowerPoint transform target contains unsupported compatibility markup',
    );
  }

  const transformName = qualifiedName(drawingPrefix, 'xfrm');
  const offsetName = qualifiedName(drawingPrefix, 'off');
  const extentName = qualifiedName(drawingPrefix, 'ext');
  const attributePattern =
    '((?:\\s+[A-Za-z_][\\w.:-]*\\s*=\\s*(?:"[^"]*"|\'[^\']*\'))*)';
  const transformPattern = new RegExp(
    `<${escapedPattern(transformName)}${attributePattern}\\s*>\\s*` +
      `<${escapedPattern(offsetName)}${attributePattern}\\s*\\/>\\s*` +
      `<${escapedPattern(extentName)}${attributePattern}\\s*\\/>\\s*` +
      `<\\/${escapedPattern(transformName)}\\s*>`,
    'g',
  );
  const matches = [...shape.matchAll(transformPattern)];
  if (matches.length !== 1) {
    unsupportedEdit(
      'PowerPoint transform target must contain one simple shape transform',
    );
  }
  const match = matches[0];
  if (match === undefined) {
    unsupportedEdit('PowerPoint shape transform disappeared');
  }
  const transformAttributes = match[1] ?? '';
  const offsetAttributes = match[2] ?? '';
  const extentAttributes = match[3] ?? '';
  const source = {
    flipHorizontal: booleanAttribute(transformAttributes, 'flipH'),
    flipVertical: booleanAttribute(transformAttributes, 'flipV'),
    height: integerAttribute(extentAttributes, 'cy'),
    rotation: optionalIntegerAttribute(transformAttributes, 'rot'),
    width: integerAttribute(extentAttributes, 'cx'),
    x: integerAttribute(offsetAttributes, 'x'),
    y: integerAttribute(offsetAttributes, 'y'),
  };
  const expected = {
    flipHorizontal: operation.expectedTransform.flipHorizontal ?? false,
    flipVertical: operation.expectedTransform.flipVertical ?? false,
    height: pointsToEmu(operation.expectedTransform.height),
    rotation: degreesToAngle(operation.expectedTransform.rotation ?? 0),
    width: pointsToEmu(operation.expectedTransform.width),
    x: pointsToEmu(operation.expectedTransform.x),
    y: pointsToEmu(operation.expectedTransform.y),
  };
  if (JSON.stringify(source) !== JSON.stringify(expected)) {
    unsupportedEdit(
      'PowerPoint transform source XML does not match its preview precondition',
    );
  }
  const replacementAttributes = [
    operation.value.rotation === 0
      ? ''
      : ` rot="${degreesToAngle(operation.value.rotation ?? 0)}"`,
    operation.value.flipHorizontal ? ' flipH="1"' : '',
    operation.value.flipVertical ? ' flipV="1"' : '',
  ].join('');
  const replacement =
    `<${transformName}${replacementAttributes}>` +
    `<${offsetName} x="${pointsToEmu(operation.value.x)}" y="${pointsToEmu(operation.value.y)}"/>` +
    `<${extentName} cx="${pointsToEmu(operation.value.width)}" cy="${pointsToEmu(operation.value.height)}"/>` +
    `</${transformName}>`;
  const matchStart = range.start + (match.index ?? 0);
  const matchEnd = matchStart + match[0].length;
  return `${xml.slice(0, matchStart)}${replacement}${xml.slice(matchEnd)}`;
}

function byteEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false;
  return left.every((value, index) => value === right[index]);
}

async function partPayloads(
  archive: JSZip,
  limits: ResolvedPptxResourceLimits,
): Promise<Map<string, Uint8Array>> {
  const result = new Map<string, Uint8Array>();
  for (const file of Object.values(archive.files)) {
    if (file.dir) continue;
    result.set(file.name, await readZipEntryBytes(file, limits.maxPartBytes));
  }
  return result;
}

function assertSafeEditablePackage(archive: JSZip): void {
  for (const name of Object.keys(archive.files)) {
    const normalized = name.toLowerCase();
    if (
      normalized.startsWith('_xmlsignatures/') ||
      normalized.endsWith('/vbaproject.bin') ||
      normalized === 'ppt/vbaproject.bin'
    ) {
      unsupportedEdit(
        'PowerPoint text edit does not modify signed or macro-enabled packages',
      );
    }
  }
}

function decodeEditableXml(
  bytes: Uint8Array,
  limits: ResolvedPptxResourceLimits,
): string {
  if (
    (bytes[0] === 0xff && bytes[1] === 0xfe) ||
    (bytes[0] === 0xfe && bytes[1] === 0xff)
  ) {
    unsupportedEdit('PowerPoint text edit requires UTF-8 slide XML');
  }
  const xml = decodeXmlBytes(bytes);
  assertXmlComplexity(xml, {
    maxDepth: limits.maxXmlDepth,
    maxNodes: limits.maxXmlNodes,
  });
  if (/encoding\s*=\s*["']utf-16["']/i.test(xml)) {
    unsupportedEdit('PowerPoint text edit requires UTF-8 slide XML');
  }
  return xml;
}

export async function patchPptxOperations(
  bytes: Uint8Array,
  document: PptxDocument,
  operations: readonly PptxRoundTripOperation[],
  limits: ResolvedPptxResourceLimits,
): Promise<PptxPatchedPackage> {
  const archive = await JSZip.loadAsync(bytes);
  assertPptxArchiveWithinLimits(archive, limits);
  assertSafeEditablePackage(archive);
  const [slides, sourcePayloads] = await Promise.all([
    resolveSlideParts(archive, limits),
    partPayloads(archive, limits),
  ]);
  if (slides.length !== document.slides.length) {
    unsupportedEdit(
      'PowerPoint text edit slide order does not match the parsed document',
    );
  }

  const patchedParts = new Set<string>();
  const editedXml = new Map<string, string>();
  for (const operation of operations) {
    const target =
      operation.kind === 'replace-text'
        ? textTarget(operation, document)
        : transformTarget(operation, document);
    const slidePart = slides[target.slideIndex];
    if (slidePart === undefined) {
      unsupportedEdit('PowerPoint text edit slide target does not exist');
    }
    const sourceBytes = sourcePayloads.get(slidePart);
    if (sourceBytes === undefined) {
      unsupportedEdit('PowerPoint text edit slide part is missing');
    }
    const current =
      editedXml.get(slidePart) ?? decodeEditableXml(sourceBytes, limits);
    const patched =
      operation.kind === 'replace-text'
        ? patchShapeText(current, target.shapeId, operation)
        : patchShapeTransform(current, target.shapeId, operation);
    editedXml.set(slidePart, patched);
    patchedParts.add(slidePart);
  }

  for (const [part, xml] of editedXml) {
    const entry = archive.file(part);
    if (entry === null)
      unsupportedEdit('PowerPoint text edit slide disappeared');
    archive.file(part, xml, {
      createFolders: false,
      date: entry.date,
    });
  }
  const output = await archive.generateAsync({
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
    platform: 'DOS',
    streamFiles: false,
    type: 'uint8array',
  });
  assertPptxInputWithinLimits(output, limits);

  const outputArchive = await JSZip.loadAsync(output);
  assertPptxArchiveWithinLimits(outputArchive, limits);
  const outputPayloads = await partPayloads(outputArchive, limits);
  if (
    outputPayloads.size !== sourcePayloads.size ||
    [...sourcePayloads.keys()].some((name) => !outputPayloads.has(name))
  ) {
    throw new PptxWriteError(
      'verification-failed',
      'PowerPoint text edit changed the package part inventory',
    );
  }
  for (const [name, source] of sourcePayloads) {
    const result = outputPayloads.get(name);
    if (result === undefined) {
      throw new PptxWriteError(
        'verification-failed',
        `PowerPoint text edit removed package part ${name}`,
      );
    }
    const equal = byteEqual(source, result);
    if (patchedParts.has(name) ? equal : !equal) {
      throw new PptxWriteError(
        'verification-failed',
        patchedParts.has(name)
          ? `PowerPoint text edit did not change dirty part ${name}`
          : `PowerPoint text edit changed untouched part ${name}`,
      );
    }
  }

  return {
    copiedPartCount: sourcePayloads.size - patchedParts.size,
    data: output,
    patchedPartCount: patchedParts.size,
  };
}

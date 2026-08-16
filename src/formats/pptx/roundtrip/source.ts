import { readZipEntryBytes } from '../../../common/archive/read-entry';
import {
  assertXmlComplexity,
  decodeXmlBytes,
} from '../../../common/xml/validate';
import type { PptxInput } from '../types';
import {
  assertPptxArchiveWithinLimits,
  assertPptxInputWithinLimits,
  type ResolvedPptxResourceLimits,
} from '../internal/resource-limits';
import type { PptxRoundTripConformance } from './types';
import { sha256Bytes } from './digest';
import JSZip from 'jszip';
import { parse as parseXml } from 'txml/txml';

const PRESENTATION_PART = 'ppt/presentation.xml';
const STRICT_PRESENTATION_NAMESPACE =
  'http://purl.oclc.org/ooxml/presentationml/main';
const TRANSITIONAL_PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';

export interface NormalizedPptxRoundTripInput {
  byteLength: number;
  bytes: Uint8Array;
  data: Blob | Uint8Array;
  sha256: string;
}

export interface PptxRoundTripPackageInspection {
  conformance: PptxRoundTripConformance;
  partCount: number;
}

interface ParsedXmlNode {
  attributes: object;
  children: Array<ParsedXmlNode | string>;
  tagName: string;
}

function cloneByteInput(input: ArrayBuffer | Uint8Array): Uint8Array {
  return input instanceof ArrayBuffer
    ? new Uint8Array(input.slice(0))
    : new Uint8Array(input);
}

export async function normalizePptxRoundTripInput(
  input: PptxInput,
  limits: ResolvedPptxResourceLimits,
): Promise<NormalizedPptxRoundTripInput> {
  assertPptxInputWithinLimits(input, limits);
  if (input instanceof Blob) {
    const data = input.slice(0, input.size, input.type);
    const bytes = new Uint8Array(await data.arrayBuffer());
    return {
      byteLength: bytes.byteLength,
      bytes,
      data,
      sha256: await sha256Bytes(bytes),
    };
  }

  const bytes = cloneByteInput(input);
  return {
    byteLength: bytes.byteLength,
    bytes,
    data: bytes,
    sha256: await sha256Bytes(bytes),
  };
}

function firstElement(
  nodes: Array<ParsedXmlNode | string>,
): ParsedXmlNode | undefined {
  return nodes.find(
    (node): node is ParsedXmlNode =>
      typeof node !== 'string' && node.tagName !== '?xml',
  );
}

function rootNamespace(root: ParsedXmlNode): string | undefined {
  const separator = root.tagName.indexOf(':');
  const prefix = separator === -1 ? '' : root.tagName.slice(0, separator);
  const attributeName = prefix === '' ? 'xmlns' : `xmlns:${prefix}`;
  const value = (root.attributes as Record<string, unknown>)[attributeName];
  return typeof value === 'string' ? value : undefined;
}

function localName(qualifiedName: string): string {
  return qualifiedName.slice(qualifiedName.indexOf(':') + 1);
}

export async function inspectPptxRoundTripPackage(
  bytes: Uint8Array,
  limits: ResolvedPptxResourceLimits,
): Promise<PptxRoundTripPackageInspection> {
  assertPptxInputWithinLimits(bytes, limits);
  const archive = await JSZip.loadAsync(bytes);
  assertPptxArchiveWithinLimits(archive, limits);
  const partCount = Object.values(archive.files).filter(
    (part) => !part.dir,
  ).length;
  const presentation = archive.file(PRESENTATION_PART);
  if (presentation === null) return { conformance: 'unknown', partCount };

  const xmlBytes = await readZipEntryBytes(presentation, limits.maxXmlBytes);
  const xml = decodeXmlBytes(xmlBytes);
  assertXmlComplexity(xml, {
    maxDepth: limits.maxXmlDepth,
    maxNodes: limits.maxXmlNodes,
  });
  const root = firstElement(parseXml(xml));
  if (root === undefined || localName(root.tagName) !== 'presentation') {
    return { conformance: 'unknown', partCount };
  }

  let conformance: PptxRoundTripConformance;
  switch (rootNamespace(root)) {
    case STRICT_PRESENTATION_NAMESPACE:
      conformance = 'strict';
      break;
    case TRANSITIONAL_PRESENTATION_NAMESPACE:
      conformance = 'transitional';
      break;
    default:
      conformance = 'unknown';
  }
  return { conformance, partCount };
}

export async function detectPptxRoundTripConformance(
  bytes: Uint8Array,
  limits: ResolvedPptxResourceLimits,
): Promise<PptxRoundTripConformance> {
  const inspection = await inspectPptxRoundTripPackage(bytes, limits);
  return inspection.conformance;
}

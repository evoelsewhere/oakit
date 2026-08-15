import type JSZip from 'jszip';
import { parse } from 'txml';

import {
  readZipEntryBytes,
  ZipExpansionBudgetLimitError,
  ZipEntrySizeLimitError,
} from '../archive/read-entry';

export type XmlValue =
  XmlNode | XmlValue[] | string | number | boolean | null | undefined;

export interface XmlNode {
  [key: string]: XmlValue;
}

export type XmlReadResult<T> =
  | { status: 'ok'; value: T }
  | { status: 'missing' }
  | { status: 'error'; error: unknown; phase: 'limit' | 'parse' | 'read' };

export interface XmlReadLimits {
  consumeBytes?: (byteLength: number) => void;
  consumeNodes?: (nodeCount: number) => void;
  maxBytes?: number;
  maxDepth?: number;
  maxNodes?: number;
}

export class XmlComplexityLimitError extends Error {
  readonly actual: number;
  readonly limit: number;
  readonly limitName: 'maxTotalXmlNodes' | 'maxXmlDepth' | 'maxXmlNodes';

  constructor(
    limitName: 'maxTotalXmlNodes' | 'maxXmlDepth' | 'maxXmlNodes',
    actual: number,
    limit: number,
  ) {
    super(`XML resource limit ${limitName} exceeded: ${actual} > ${limit}`);
    this.name = 'XmlComplexityLimitError';
    this.actual = actual;
    this.limit = limit;
    this.limitName = limitName;
  }
}

export class XmlStructureError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'XmlStructureError';
  }
}

interface TxmlNode {
  attributes?: Record<string, string>;
  children?: Array<TxmlNode | string>;
  tagName?: string;
}

interface SimplifyState {
  documentOrder: number;
}

const CANONICAL_NAMESPACE_PREFIXES: Readonly<Record<string, string>> = {
  'http://purl.oclc.org/ooxml/drawingml/chart': 'c',
  'http://purl.oclc.org/ooxml/drawingml/diagram': 'dgm',
  'http://purl.oclc.org/ooxml/drawingml/main': 'a',
  'http://purl.oclc.org/ooxml/officeDocument/math': 'm',
  'http://purl.oclc.org/ooxml/officeDocument/relationships': 'r',
  'http://purl.oclc.org/ooxml/presentationml/main': 'p',
  'http://schemas.microsoft.com/office/drawing/2008/diagram': 'dsp',
  'http://schemas.microsoft.com/office/drawing/2010/main': 'a14',
  'http://schemas.openxmlformats.org/drawingml/2006/chart': 'c',
  'http://schemas.openxmlformats.org/drawingml/2006/diagram': 'dgm',
  'http://schemas.openxmlformats.org/drawingml/2006/main': 'a',
  'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
  'http://schemas.openxmlformats.org/officeDocument/2006/math': 'm',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships': 'r',
  'http://schemas.openxmlformats.org/package/2006/content-types': '',
  'http://schemas.openxmlformats.org/package/2006/relationships': '',
  'http://schemas.openxmlformats.org/presentationml/2006/main': 'p',
};

const RESERVED_CANONICAL_PREFIXES = new Set(
  Object.values(CANONICAL_NAMESPACE_PREFIXES).filter(Boolean),
);

function extendNamespaceBindings(
  parent: ReadonlyMap<string, string>,
  attributes: Readonly<Record<string, string>>,
): Map<string, string> {
  const bindings = new Map(parent);
  for (const [name, value] of Object.entries(attributes)) {
    if (name === 'xmlns') bindings.set('', value);
    else if (name.startsWith('xmlns:')) bindings.set(name.slice(6), value);
  }
  return bindings;
}

function normalizeQualifiedName(
  name: string,
  bindings: ReadonlyMap<string, string>,
  attribute: boolean,
): string {
  if (name === 'xmlns' || name.startsWith('xmlns:')) return name;

  const separatorIndex = name.indexOf(':');
  const sourcePrefix = separatorIndex < 0 ? '' : name.slice(0, separatorIndex);
  const localName = separatorIndex < 0 ? name : name.slice(separatorIndex + 1);
  if (attribute && !sourcePrefix) return name;

  const namespace = bindings.get(sourcePrefix);
  if (!namespace) return name;
  const canonicalPrefix = CANONICAL_NAMESPACE_PREFIXES[namespace];
  if (canonicalPrefix === undefined) {
    return RESERVED_CANONICAL_PREFIXES.has(sourcePrefix)
      ? `ns_${sourcePrefix}:${localName}`
      : name;
  }
  return canonicalPrefix ? `${canonicalPrefix}:${localName}` : localName;
}

function normalizeAttributes(
  attributes: Readonly<Record<string, string>>,
  bindings: ReadonlyMap<string, string>,
): Record<string, string> {
  const normalized: Record<string, string> = {};
  for (const [name, value] of Object.entries(attributes)) {
    normalized[normalizeQualifiedName(name, bindings, true)] = value;
  }
  return normalized;
}

function isWhitespaceTextNode(node: TxmlNode | string): boolean {
  return typeof node === 'string' && node.trim() === '';
}

function simplifyLosslessWithState(
  children: Array<TxmlNode | string>,
  parentAttributes: Record<string, string>,
  state: SimplifyState,
  parentNamespaces: ReadonlyMap<string, string>,
): XmlValue {
  const output: XmlNode = {};
  if (children.length === 0) return output;

  if (children.length === 1 && typeof children[0] === 'string') {
    return Object.keys(parentAttributes).length > 0
      ? {
          attrs: { order: state.documentOrder++, ...parentAttributes },
          value: children[0],
        }
      : children[0];
  }

  for (const child of children) {
    if (isWhitespaceTextNode(child)) continue;
    if (typeof child !== 'object') return undefined;
    if (child.tagName === '?xml') continue;
    if (!child.tagName) continue;

    const namespaces = extendNamespaceBindings(
      parentNamespaces,
      child.attributes ?? {},
    );
    const tagName = normalizeQualifiedName(child.tagName, namespaces, false);
    const childAttributes = normalizeAttributes(
      child.attributes ?? {},
      namespaces,
    );
    const existing = output[tagName];
    const values = Array.isArray(existing)
      ? existing
      : existing
        ? [existing]
        : [];
    const value = simplifyLosslessWithState(
      child.children ?? [],
      childAttributes,
      state,
      namespaces,
    );

    if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
      const attrs = value.attrs;
      value.attrs = {
        order: state.documentOrder++,
        ...(typeof attrs === 'object' && attrs !== null && !Array.isArray(attrs)
          ? attrs
          : {}),
        ...childAttributes,
      };
    }
    values.push(value);
    output[tagName] = values;
  }

  for (const key of Object.keys(output)) {
    const values = output[key];
    if (Array.isArray(values) && values.length === 1) output[key] = values[0];
  }

  return output;
}

/** Convert txml's lossless tree into the object shape consumed by format parsers. */
export function simplifyLossless(
  children: Array<TxmlNode | string>,
  parentAttributes: Record<string, string> = {},
): XmlValue {
  return simplifyLosslessWithState(
    children,
    parentAttributes,
    {
      documentOrder: 0,
    },
    new Map(),
  );
}

function markupEnd(xml: string, start: number, declaration: boolean): number {
  let quote = '';
  let subsetDepth = 0;
  for (let index = start; index < xml.length; index++) {
    const character = xml[index]!;
    if (quote) {
      if (character === quote) quote = '';
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
      continue;
    }
    if (declaration) {
      if (character === '[') subsetDepth++;
      else if (character === ']' && subsetDepth > 0) subsetDepth--;
    }
    if (character === '>' && subsetDepth === 0) return index;
  }
  return xml.length - 1;
}

/** Reject pathological nesting before the recursive XML parser sees it. */
export function assertXmlComplexity(
  xml: string,
  limits: Pick<XmlReadLimits, 'maxDepth' | 'maxNodes'>,
): number {
  let depth = 0;
  let nodes = 0;
  let index = 0;
  const openElements: string[] = [];

  while (index < xml.length) {
    const opening = xml.indexOf('<', index);
    if (opening < 0) break;

    if (xml.startsWith('<!--', opening)) {
      const end = xml.indexOf('-->', opening + 4);
      index = end < 0 ? xml.length : end + 3;
      continue;
    }
    if (xml.startsWith('<![CDATA[', opening)) {
      const end = xml.indexOf(']]>', opening + 9);
      index = end < 0 ? xml.length : end + 3;
      continue;
    }
    if (xml.startsWith('<?', opening)) {
      const end = xml.indexOf('?>', opening + 2);
      index = end < 0 ? xml.length : end + 2;
      continue;
    }
    if (xml.startsWith('<!', opening)) {
      index = markupEnd(xml, opening + 2, true) + 1;
      continue;
    }

    const end = markupEnd(xml, opening + 1, false);
    if (xml.startsWith('</', opening)) {
      const closingName = xml.slice(opening + 2, end).trim();
      const expectedName = openElements.pop();
      if (!expectedName || closingName !== expectedName) {
        throw new XmlStructureError(
          `Unexpected XML closing tag ${closingName || '(empty)'}; expected ${expectedName ?? '(none)'}`,
        );
      }
      depth--;
    } else {
      const tagContent = xml.slice(opening + 1, end).trimStart();
      const elementName = /^[^\s/>]+/.exec(tagContent)?.[0];
      if (!elementName) {
        throw new XmlStructureError('XML opening tag has no element name');
      }
      nodes++;
      if (limits.maxNodes !== undefined && nodes > limits.maxNodes) {
        throw new XmlComplexityLimitError(
          'maxXmlNodes',
          nodes,
          limits.maxNodes,
        );
      }
      const nodeDepth = depth + 1;
      if (limits.maxDepth !== undefined && nodeDepth > limits.maxDepth) {
        throw new XmlComplexityLimitError(
          'maxXmlDepth',
          nodeDepth,
          limits.maxDepth,
        );
      }
      const selfClosing = xml
        .slice(opening + 1, end)
        .trimEnd()
        .endsWith('/');
      if (!selfClosing) {
        depth = nodeDepth;
        openElements.push(elementName);
      }
    }
    index = end + 1;
  }
  if (openElements.length > 0) {
    throw new XmlStructureError(
      `Unclosed XML element ${openElements.at(-1) ?? '(unknown)'}`,
    );
  }
  return nodes;
}

/** Read an XML part while preserving missing and failed states for callers. */
export async function readXmlFileResult<T extends XmlValue = XmlValue>(
  zip: JSZip,
  filename: string,
  limits: XmlReadLimits = {},
): Promise<XmlReadResult<T>> {
  if (!filename) return { status: 'missing' };
  const file = zip.file(filename);
  if (!file) return { status: 'missing' };

  let data: string;
  try {
    if (limits.maxBytes === undefined) {
      data = await file.async('string');
    } else {
      const bytes = await readZipEntryBytes(
        file,
        limits.maxBytes,
        limits.consumeBytes,
      );
      data = new TextDecoder().decode(bytes);
    }
  } catch (error) {
    return {
      status: 'error',
      phase:
        error instanceof ZipEntrySizeLimitError ||
        error instanceof ZipExpansionBudgetLimitError
          ? 'limit'
          : 'read',
      error,
    };
  }

  try {
    const nodeCount = assertXmlComplexity(data, limits);
    limits.consumeNodes?.(nodeCount);
    return {
      status: 'ok',
      value: simplifyLossless(parse(data, { keepWhitespace: true })) as T,
    };
  } catch (error) {
    return {
      status: 'error',
      phase: error instanceof XmlComplexityLimitError ? 'limit' : 'parse',
      error,
    };
  }
}

/** Read and simplify an XML part. Missing or invalid optional parts return null. */
export async function readXmlFile<T extends XmlValue = XmlValue>(
  zip: JSZip,
  filename: string,
): Promise<T | null> {
  const result = await readXmlFileResult<T>(zip, filename);
  return result.status === 'ok' ? result.value : null;
}

/** @deprecated Upstream misspelling retained for compatibility. */
export const simplifyLostLess = simplifyLossless;

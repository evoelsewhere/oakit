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
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
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
    const normalizedName = normalizeQualifiedName(name, bindings, true);
    if (
      name !== 'xmlns' &&
      !name.startsWith('xmlns:') &&
      Object.hasOwn(normalized, normalizedName)
    ) {
      throw new XmlStructureError(
        `XML element has duplicate expanded attribute ${normalizedName}`,
      );
    }
    normalized[normalizedName] = value;
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

function decodeXmlBytes(bytes: Uint8Array): string {
  let encoding = 'utf-8';
  if (bytes[0] === 0xff && bytes[1] === 0xfe) encoding = 'utf-16le';
  else if (bytes[0] === 0xfe && bytes[1] === 0xff) encoding = 'utf-16be';

  try {
    return new TextDecoder(encoding, { fatal: true }).decode(bytes);
  } catch (cause) {
    throw new XmlStructureError(`Invalid ${encoding.toUpperCase()} XML`, {
      cause,
    });
  }
}

function isValidXmlCodePoint(codePoint: number): boolean {
  return (
    codePoint === 0x9 ||
    codePoint === 0xa ||
    codePoint === 0xd ||
    (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
    (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
    (codePoint >= 0x10_000 && codePoint <= 0x10_ffff)
  );
}

function assertValidXmlCharacters(value: string): void {
  for (const character of value) {
    if (!isValidXmlCodePoint(character.codePointAt(0)!)) {
      throw new XmlStructureError('XML contains an invalid character');
    }
  }
}

function assertValidEntityReferences(value: string): void {
  let index = value.indexOf('&');
  while (index >= 0) {
    const end = value.indexOf(';', index + 1);
    if (end < 0) {
      throw new XmlStructureError(
        'XML entity reference is missing a semicolon',
      );
    }
    const entity = value.slice(index + 1, end);
    if (!/^(?:amp|apos|gt|lt|quot)$/.test(entity)) {
      const decimal = /^#(\d+)$/.exec(entity)?.[1];
      const hexadecimal = /^#x([\da-f]+)$/i.exec(entity)?.[1];
      const codePoint = decimal
        ? Number.parseInt(decimal, 10)
        : hexadecimal
          ? Number.parseInt(hexadecimal, 16)
          : Number.NaN;
      if (!Number.isInteger(codePoint) || !isValidXmlCodePoint(codePoint)) {
        throw new XmlStructureError(`Invalid XML entity reference &${entity};`);
      }
    }
    index = value.indexOf('&', end + 1);
  }
}

interface StartTagDetails {
  elementName: string;
  selfClosing: boolean;
}

function inspectStartTag(tagContent: string): StartTagDetails {
  let contentEnd = tagContent.length;
  while (contentEnd > 0 && /\s/.test(tagContent[contentEnd - 1]!)) {
    contentEnd--;
  }

  const selfClosing = tagContent[contentEnd - 1] === '/';
  if (selfClosing) {
    contentEnd--;
    while (contentEnd > 0 && /\s/.test(tagContent[contentEnd - 1]!)) {
      contentEnd--;
    }
  }

  const content = tagContent.slice(0, contentEnd);
  const elementName = /^[A-Za-z_:][A-Za-z\d_.:-]*/.exec(content)?.[0];
  if (!elementName) {
    throw new XmlStructureError('XML opening tag has no valid element name');
  }

  let cursor = elementName.length;
  const attributes = new Set<string>();
  while (cursor < content.length) {
    if (!/\s/.test(content[cursor]!)) {
      throw new XmlStructureError(
        `XML attribute after ${elementName} is not separated by whitespace`,
      );
    }
    while (cursor < content.length && /\s/.test(content[cursor]!)) cursor++;
    if (cursor >= content.length) break;

    const attributeName = /^[A-Za-z_:][A-Za-z\d_.:-]*/.exec(
      content.slice(cursor),
    )?.[0];
    if (!attributeName) {
      throw new XmlStructureError(
        `XML element ${elementName} has an invalid attribute name`,
      );
    }
    if (attributes.has(attributeName)) {
      throw new XmlStructureError(
        `XML element ${elementName} has duplicate attribute ${attributeName}`,
      );
    }
    attributes.add(attributeName);
    cursor += attributeName.length;

    while (cursor < content.length && /\s/.test(content[cursor]!)) cursor++;
    if (content[cursor] !== '=') {
      throw new XmlStructureError(
        `XML attribute ${attributeName} must have a value`,
      );
    }
    cursor++;
    while (cursor < content.length && /\s/.test(content[cursor]!)) cursor++;

    const quote = content[cursor];
    if (quote !== '"' && quote !== "'") {
      throw new XmlStructureError(
        `XML attribute ${attributeName} must use a quoted value`,
      );
    }
    const valueEnd = content.indexOf(quote, cursor + 1);
    if (valueEnd < 0) {
      throw new XmlStructureError(
        `XML attribute ${attributeName} has an unclosed value`,
      );
    }
    if (content.slice(cursor + 1, valueEnd).includes('<')) {
      throw new XmlStructureError(
        `XML attribute ${attributeName} contains an invalid character`,
      );
    }
    cursor = valueEnd + 1;
  }

  return { elementName, selfClosing };
}

/** Reject pathological nesting before the recursive XML parser sees it. */
export function assertXmlComplexity(
  xml: string,
  limits: Pick<XmlReadLimits, 'maxDepth' | 'maxNodes'>,
): number {
  let depth = 0;
  let nodes = 0;
  let index = 0;
  let rootElements = 0;
  const openElements: string[] = [];

  assertValidXmlCharacters(xml);

  while (index < xml.length) {
    const opening = xml.indexOf('<', index);
    const textEnd = opening < 0 ? xml.length : opening;
    const text = xml.slice(index, textEnd);
    assertValidEntityReferences(text);
    if (text.includes(']]>')) {
      throw new XmlStructureError(
        'XML CDATA terminator appears outside a CDATA section',
      );
    }
    if (depth === 0 && text.trim()) {
      throw new XmlStructureError('XML text is not inside the document root');
    }
    if (opening < 0) break;

    if (xml.startsWith('<!--', opening)) {
      const end = xml.indexOf('-->', opening + 4);
      if (end < 0) throw new XmlStructureError('Unclosed XML comment');
      if (xml.slice(opening + 4, end).includes('--')) {
        throw new XmlStructureError('XML comment contains a double hyphen');
      }
      index = end < 0 ? xml.length : end + 3;
      continue;
    }
    if (xml.startsWith('<![CDATA[', opening)) {
      const end = xml.indexOf(']]>', opening + 9);
      if (end < 0) throw new XmlStructureError('Unclosed XML CDATA section');
      const content = xml.slice(opening + 9, end);
      if (depth === 0 && content.trim()) {
        throw new XmlStructureError(
          'XML CDATA is not inside the document root',
        );
      }
      index = end < 0 ? xml.length : end + 3;
      continue;
    }
    if (xml.startsWith('<?', opening)) {
      const end = xml.indexOf('?>', opening + 2);
      if (end < 0) {
        throw new XmlStructureError('Unclosed XML processing instruction');
      }
      index = end < 0 ? xml.length : end + 2;
      continue;
    }
    if (xml.startsWith('<!', opening)) {
      throw new XmlStructureError('XML declarations with <! are not allowed');
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
      assertValidEntityReferences(tagContent);
      const { elementName, selfClosing } = inspectStartTag(tagContent);
      nodes++;
      if (limits.maxNodes !== undefined && nodes > limits.maxNodes) {
        throw new XmlComplexityLimitError(
          'maxXmlNodes',
          nodes,
          limits.maxNodes,
        );
      }
      const nodeDepth = depth + 1;
      if (depth === 0) {
        rootElements++;
        if (rootElements > 1) {
          throw new XmlStructureError(
            'XML document has multiple root elements',
          );
        }
      }
      if (limits.maxDepth !== undefined && nodeDepth > limits.maxDepth) {
        throw new XmlComplexityLimitError(
          'maxXmlDepth',
          nodeDepth,
          limits.maxDepth,
        );
      }
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
  if (rootElements !== 1) {
    throw new XmlStructureError('XML document must contain one root element');
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

  let bytes: Uint8Array;
  try {
    bytes = await readZipEntryBytes(
      file,
      limits.maxBytes ?? Number.MAX_SAFE_INTEGER,
      limits.consumeBytes,
    );
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
    const data = decodeXmlBytes(bytes);
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

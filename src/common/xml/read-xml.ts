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

function isWhitespaceTextNode(node: TxmlNode | string): boolean {
  return typeof node === 'string' && node.trim() === '';
}

function simplifyLosslessWithState(
  children: Array<TxmlNode | string>,
  parentAttributes: Record<string, string>,
  state: SimplifyState,
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

    const existing = output[child.tagName];
    const values = Array.isArray(existing)
      ? existing
      : existing
        ? [existing]
        : [];
    const value = simplifyLosslessWithState(
      child.children ?? [],
      child.attributes ?? {},
      state,
    );

    if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
      const attrs = value.attrs;
      value.attrs = {
        order: state.documentOrder++,
        ...(typeof attrs === 'object' && attrs !== null && !Array.isArray(attrs)
          ? attrs
          : {}),
        ...(child.attributes ?? {}),
      };
    }
    values.push(value);
    output[child.tagName] = values;
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
  return simplifyLosslessWithState(children, parentAttributes, {
    documentOrder: 0,
  });
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

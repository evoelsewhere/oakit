import type JSZip from 'jszip';
import { parse } from 'txml';

export type XmlValue =
  XmlNode | XmlValue[] | string | number | boolean | null | undefined;

export interface XmlNode {
  [key: string]: XmlValue;
}

export type XmlReadResult<T> =
  | { status: 'ok'; value: T }
  | { status: 'missing' }
  | { status: 'error'; error: unknown; phase: 'parse' | 'read' };

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

/** Read an XML part while preserving missing and failed states for callers. */
export async function readXmlFileResult<T extends XmlValue = XmlValue>(
  zip: JSZip,
  filename: string,
): Promise<XmlReadResult<T>> {
  if (!filename) return { status: 'missing' };
  const file = zip.file(filename);
  if (!file) return { status: 'missing' };

  let data: string;
  try {
    data = await file.async('string');
  } catch (error) {
    return { status: 'error', phase: 'read', error };
  }

  try {
    return {
      status: 'ok',
      value: simplifyLossless(parse(data, { keepWhitespace: true })) as T,
    };
  } catch (error) {
    return { status: 'error', phase: 'parse', error };
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

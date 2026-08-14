import type JSZip from 'jszip';
import { parse } from 'txml';

export type XmlValue =
  XmlNode | XmlValue[] | string | number | boolean | null | undefined;

export interface XmlNode {
  [key: string]: XmlValue;
}

interface TxmlNode {
  attributes?: Record<string, string>;
  children?: Array<TxmlNode | string>;
  tagName?: string;
}

let documentOrder = 0;

function isWhitespaceTextNode(node: TxmlNode | string): boolean {
  return typeof node === 'string' && node.trim() === '';
}

/** Convert txml's lossless tree into the object shape consumed by format parsers. */
export function simplifyLossless(
  children: Array<TxmlNode | string>,
  parentAttributes: Record<string, string> = {},
): XmlValue {
  const output: XmlNode = {};
  if (children.length === 0) return output;

  if (children.length === 1 && typeof children[0] === 'string') {
    return Object.keys(parentAttributes).length > 0
      ? {
          attrs: { order: documentOrder++, ...parentAttributes },
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
    const value = simplifyLossless(
      child.children ?? [],
      child.attributes ?? {},
    );

    if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
      const attrs = value.attrs;
      value.attrs = {
        order: documentOrder++,
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

/** Read and simplify an XML part. Missing optional parts return null. */
export async function readXmlFile<T extends XmlValue = XmlValue>(
  zip: JSZip,
  filename: string,
): Promise<T | null> {
  if (!filename) return null;
  try {
    const file = zip.file(filename);
    if (!file) return null;

    const data = await file.async('string');
    return simplifyLossless(parse(data, { keepWhitespace: true })) as T;
  } catch {
    return null;
  }
}

/** @deprecated Upstream misspelling retained for compatibility. */
export const simplifyLostLess = simplifyLossless;

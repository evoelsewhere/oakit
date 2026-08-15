import { type XmlNode, XmlStructureError, type XmlValue } from './types';

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
  const separatorIndex = name.indexOf(':');
  const hasPrefix = name.includes(':');
  const sourcePrefix = hasPrefix ? name.slice(0, separatorIndex) : '';
  const localName = hasPrefix ? name.slice(separatorIndex + 1) : name;
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
    if (Object.hasOwn(normalized, normalizedName)) {
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
    const values = Array.isArray(existing) ? existing : [];
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
        ...(typeof attrs === 'object' && !Array.isArray(attrs) ? attrs : {}),
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
    { documentOrder: 0 },
    new Map(),
  );
}

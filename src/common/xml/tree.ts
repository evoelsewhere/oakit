type XmlContainer = Readonly<Record<string, unknown>>;

const xmlNodeOrders = new WeakMap<object, number>();

/**
 * Value returned by the compatibility path lookup.
 *
 * OOXML elements are dynamically addressed and may be consumed as a node,
 * list, or scalar depending on the schema path. The intersection models that
 * legacy contract without leaking `any`; new schema-aware code should pass an
 * explicit generic result type instead.
 */
interface XmlLookupRecord {
  [key: string]: XmlLookupValue;
}

type XmlLookupArray = XmlLookupValue[];

export type XmlLookupValue = string & XmlLookupArray & XmlLookupRecord;

function isContainer(value: unknown): value is XmlContainer {
  return typeof value === 'object' && value !== null;
}

/** Attach parser-only document order without occupying an XML attribute name. */
export function setXmlNodeOrder(node: object, order: number): void {
  xmlNodeOrders.set(node, order);
}

/**
 * Read parser-only document order.
 *
 * The attribute fallback keeps manually assembled compatibility trees working;
 * normalized XML always has WeakMap metadata, which takes precedence over an
 * authored attribute named `order`.
 */
export function getXmlNodeOrder(node: unknown): number | undefined {
  if (!isContainer(node)) return undefined;
  const metadataOrder = xmlNodeOrders.get(node);
  if (metadataOrder !== undefined) return metadataOrder;

  const attrs = node.attrs;
  if (!isContainer(attrs)) return undefined;
  const legacyOrder = attrs.order;
  if (typeof legacyOrder === 'number') {
    return Number.isSafeInteger(legacyOrder) && legacyOrder >= 0
      ? legacyOrder
      : undefined;
  }
  if (typeof legacyOrder !== 'string' || !/^(0|[1-9]\d*)$/.test(legacyOrder)) {
    return undefined;
  }
  const parsed = Number(legacyOrder);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}

export function getTextByPathList(
  node: unknown,
  path: readonly string[],
): XmlLookupValue;
export function getTextByPathList<T>(
  node: unknown,
  path: readonly string[],
): T | undefined;
export function getTextByPathList(
  node: unknown,
  path: readonly string[],
): unknown {
  let current = node;

  for (const key of path) {
    if (!isContainer(current)) return undefined;
    current = current[key];
    if (current === null) return undefined;
  }

  return current;
}

export function eachElement(
  node: unknown,
  callback: (
    item: unknown,
    index: number,
  ) => string | number | boolean | null | undefined,
): string | null | undefined {
  if (node === null) return null;
  if (node === undefined) return undefined;
  const items: readonly unknown[] = Array.isArray(node) ? node : [node];
  return items.map((item, index) => String(callback(item, index))).join('');
}

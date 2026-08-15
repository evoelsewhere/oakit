type XmlContainer = Readonly<Record<string, unknown>>;

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

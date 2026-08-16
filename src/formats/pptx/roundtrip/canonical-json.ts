const MAX_CANONICAL_JSON_DEPTH = 64;

function canonicalValue(
  value: unknown,
  active: WeakSet<object>,
  depth: number,
): string {
  if (depth > MAX_CANONICAL_JSON_DEPTH) {
    throw new TypeError('Canonical JSON exceeds the maximum depth of 64');
  }
  if (value === null) return 'null';
  switch (typeof value) {
    case 'boolean':
    case 'string':
      return JSON.stringify(value);
    case 'number':
      if (!Number.isFinite(value)) {
        throw new TypeError('Canonical JSON requires finite numbers');
      }
      return JSON.stringify(value);
    case 'object': {
      if (active.has(value)) {
        throw new TypeError('Canonical JSON does not support cycles');
      }
      active.add(value);
      try {
        if (Array.isArray(value)) {
          return `[${value.map((child) => canonicalValue(child, active, depth + 1)).join(',')}]`;
        }
        if (Object.getPrototypeOf(value) !== Object.prototype) {
          throw new TypeError('Canonical JSON requires plain objects');
        }
        const record = value as Record<string, unknown>;
        const properties = Object.keys(record)
          .sort()
          .map(
            (key) =>
              `${JSON.stringify(key)}:${canonicalValue(record[key], active, depth + 1)}`,
          );
        return `{${properties.join(',')}}`;
      } finally {
        active.delete(value);
      }
    }
    default:
      throw new TypeError(`Canonical JSON does not support ${typeof value}`);
  }
}

export function canonicalJson(value: unknown): string {
  return canonicalValue(value, new WeakSet<object>(), 0);
}

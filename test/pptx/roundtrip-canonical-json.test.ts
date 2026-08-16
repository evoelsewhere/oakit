import { describe, expect, it } from 'vitest';

import { canonicalJson } from '../../src/formats/pptx/roundtrip/canonical-json';

describe('PowerPoint round-trip canonical JSON', () => {
  it('sorts object keys while preserving array order', () => {
    expect(
      canonicalJson({ z: 1, a: ['second', 'first'], nested: { y: 2, x: 1 } }),
    ).toBe('{"a":["second","first"],"nested":{"x":1,"y":2},"z":1}');
  });

  it.each([
    [null, 'null'],
    [false, 'false'],
    [true, 'true'],
    [-0, '0'],
    [1.5, '1.5'],
    ['<&', '"<&"'],
    [[], '[]'],
    [{}, '{}'],
  ])('serializes JSON value %j canonically', (value, expected) => {
    expect(canonicalJson(value)).toBe(expected);
  });

  it.each([Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY])(
    'rejects non-finite number %s',
    (value) => {
      expect(() => canonicalJson(value)).toThrow(
        new TypeError('Canonical JSON requires finite numbers'),
      );
    },
  );

  it.each([undefined, Symbol('value'), () => undefined])(
    'rejects unsupported value %s',
    (value) => {
      expect(() => canonicalJson(value)).toThrow(
        new TypeError(`Canonical JSON does not support ${typeof value}`),
      );
    },
  );

  it('rejects non-plain object instances', () => {
    expect(() => canonicalJson(new Date(0))).toThrow(
      new TypeError('Canonical JSON requires plain objects'),
    );
  });

  it('rejects direct and nested cycles without retaining traversal state', () => {
    const direct: Record<string, unknown> = {};
    direct.self = direct;
    const shared = { value: 1 };

    expect(() => canonicalJson(direct)).toThrow(
      new TypeError('Canonical JSON does not support cycles'),
    );
    expect(canonicalJson({ first: shared, second: shared })).toBe(
      '{"first":{"value":1},"second":{"value":1}}',
    );
  });

  it('bounds canonical traversal depth independently of cycle detection', () => {
    let acceptedObject: Record<string, unknown> = {};
    let rejectedObject: Record<string, unknown> = {};
    let acceptedArray: unknown = [];
    let rejectedArray: unknown = [];
    for (let depth = 0; depth < 64; depth += 1) {
      acceptedObject = { child: acceptedObject };
      rejectedObject = { child: rejectedObject };
      acceptedArray = [acceptedArray];
      rejectedArray = [rejectedArray];
    }
    rejectedObject = { child: rejectedObject };
    rejectedArray = [rejectedArray];

    expect(() => canonicalJson(acceptedObject)).not.toThrow();
    expect(() => canonicalJson(acceptedArray)).not.toThrow();
    expect(() => canonicalJson(rejectedObject)).toThrow(
      new TypeError('Canonical JSON exceeds the maximum depth of 64'),
    );
    expect(() => canonicalJson(rejectedArray)).toThrow(
      new TypeError('Canonical JSON exceeds the maximum depth of 64'),
    );
  });
});

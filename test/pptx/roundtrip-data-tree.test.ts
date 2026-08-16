import { describe, expect, it } from 'vitest';

import { resolvePptxResourceLimits } from '../../src/formats/pptx/internal/resource-limits';
import { assertPptxRoundTripDataTree } from '../../src/formats/pptx/roundtrip/data-tree';

function expectInvalid(value: unknown, message: string): void {
  expect(() =>
    assertPptxRoundTripDataTree(value, resolvePptxResourceLimits()),
  ).toThrow(message);
}

describe('PowerPoint round-trip snapshot data trees', () => {
  it('accepts plain JSON data with binary only at the source data path', () => {
    expect(() =>
      assertPptxRoundTripDataTree(
        {
          enabled: true,
          nullable: null,
          numbers: [0, 1.5],
          source: { data: new Uint8Array([1, 2, 3]) },
          text: 'snapshot',
        },
        resolvePptxResourceLimits(),
      ),
    ).not.toThrow();
    expect(() =>
      assertPptxRoundTripDataTree(
        { source: { data: new Blob([new ArrayBuffer(0)]) } },
        resolvePptxResourceLimits(),
      ),
    ).not.toThrow();
  });

  it.each([
    ['undefined', () => undefined, 'contains a non-JSON value'],
    ['bigint', () => 1n, 'contains a non-JSON value'],
    ['NaN', () => Number.NaN, 'contains a non-JSON value'],
    ['infinity', () => Number.POSITIVE_INFINITY, 'contains a non-JSON value'],
    ['Date', () => new Date(0), 'requires plain objects'],
    [
      'null prototype',
      () => {
        const value = {};
        Object.setPrototypeOf(value, null);
        return value;
      },
      'requires plain objects',
    ],
    [
      'root bytes',
      () => new Uint8Array(),
      'contains binary data outside its source',
    ],
    [
      'nested Blob',
      () => ({ nested: new Blob() }),
      'contains binary data outside its source',
    ],
  ])('rejects unsupported %s', (_name, create, message) => {
    expectInvalid(create(), message);
  });

  it('rejects symbol keys without invoking their values', () => {
    const value = { safe: true };
    Object.defineProperty(value, Symbol('hidden'), {
      enumerable: true,
      value: 'secret',
    });

    expectInvalid(value, 'contains a symbol key');
  });

  it.each([
    [
      'accessor',
      () => {
        const value = {};
        Object.defineProperty(value, 'unsafe', {
          enumerable: true,
          get: () => {
            throw new Error('must not execute');
          },
        });
        return value;
      },
      'contains an accessor or hidden property',
    ],
    [
      'hidden',
      () => {
        const value = {};
        Object.defineProperty(value, 'hidden', {
          enumerable: false,
          value: true,
        });
        return value;
      },
      'contains an accessor or hidden property',
    ],
  ])('rejects an object %s property', (_name, create, message) => {
    expectInvalid(create(), message);
  });

  it('rejects direct cycles and shared object identity', () => {
    const cycle: Record<string, unknown> = {};
    cycle.self = cycle;
    const shared = { value: 1 };

    expectInvalid(cycle, 'contains a repeated object reference');
    expectInvalid(
      { first: shared, second: shared },
      'contains a repeated object reference',
    );
  });

  it('rejects sparse, extended, accessor, and subclassed arrays', () => {
    const sparse = new Array<unknown>(1);
    const extended: unknown[] & { extra?: boolean } = [];
    extended.extra = true;
    const accessor: unknown[] = [1];
    Object.defineProperty(accessor, '0', {
      enumerable: true,
      get: () => 1,
    });
    class ArraySubclass extends Array<unknown> {}

    expectInvalid(sparse, 'contains a sparse or extended array');
    expectInvalid(extended, 'contains a sparse or extended array');
    expectInvalid(accessor, 'contains a sparse or accessor array');
    expectInvalid(new ArraySubclass(), 'requires plain arrays');
  });

  it('rejects symbol keys on arrays', () => {
    const value: unknown[] = [];
    Object.defineProperty(value, Symbol('hidden'), {
      enumerable: true,
      value: true,
    });

    expectInvalid(value, 'contains a symbol key');
  });

  it('bounds array length before inspecting sparse indexes', () => {
    expect(() =>
      assertPptxRoundTripDataTree(
        [null],
        resolvePptxResourceLimits({ maxTotalXmlNodes: 1 }),
      ),
    ).not.toThrow();
    expect(() =>
      assertPptxRoundTripDataTree(
        new Array<unknown>(2),
        resolvePptxResourceLimits({ maxTotalXmlNodes: 1 }),
      ),
    ).toThrow('exceeds the container budget');
  });

  it('bounds total containers and string code units', () => {
    expect(() =>
      assertPptxRoundTripDataTree(
        {},
        resolvePptxResourceLimits({ maxTotalXmlNodes: 1 }),
      ),
    ).not.toThrow();
    expect(() =>
      assertPptxRoundTripDataTree(
        { nested: {} },
        resolvePptxResourceLimits({ maxTotalXmlNodes: 1 }),
      ),
    ).toThrow('exceeds the container budget');
    expect(() =>
      assertPptxRoundTripDataTree(
        { a: 'b' },
        resolvePptxResourceLimits({ maxTotalUncompressedBytes: 1 }),
      ),
    ).toThrow('exceeds the string budget');
  });

  it('counts object property names against the string budget', () => {
    expect(() =>
      assertPptxRoundTripDataTree(
        { a: null },
        resolvePptxResourceLimits({ maxTotalUncompressedBytes: 1 }),
      ),
    ).not.toThrow();
    expect(() =>
      assertPptxRoundTripDataTree(
        { ab: null },
        resolvePptxResourceLimits({ maxTotalUncompressedBytes: 1 }),
      ),
    ).toThrow('exceeds the string budget');
  });

  it('accepts depth 64 and rejects depth 65', () => {
    let accepted: Record<string, unknown> = {};
    let rejected: Record<string, unknown> = {};
    for (let depth = 0; depth < 64; depth += 1) {
      accepted = { child: accepted };
      rejected = { child: rejected };
    }
    rejected = { child: rejected };

    expect(() =>
      assertPptxRoundTripDataTree(
        accepted,
        resolvePptxResourceLimits({
          maxTotalUncompressedBytes: 1_000,
          maxTotalXmlNodes: 1_000,
        }),
      ),
    ).not.toThrow();
    expect(() =>
      assertPptxRoundTripDataTree(
        rejected,
        resolvePptxResourceLimits({
          maxTotalUncompressedBytes: 1_000,
          maxTotalXmlNodes: 1_000,
        }),
      ),
    ).toThrow('exceeds the maximum data depth');
  });

  it('accepts array depth 64 and rejects array depth 65', () => {
    let accepted: unknown = [];
    let rejected: unknown = [];
    for (let depth = 0; depth < 64; depth += 1) {
      accepted = [accepted];
      rejected = [rejected];
    }
    rejected = [rejected];

    expect(() =>
      assertPptxRoundTripDataTree(
        accepted,
        resolvePptxResourceLimits({ maxTotalXmlNodes: 1_000 }),
      ),
    ).not.toThrow();
    expect(() =>
      assertPptxRoundTripDataTree(
        rejected,
        resolvePptxResourceLimits({ maxTotalXmlNodes: 1_000 }),
      ),
    ).toThrow('exceeds the maximum data depth');
  });

  it('does not treat source-shaped objects inside arrays as source data', () => {
    expectInvalid(
      [{ source: { data: new Uint8Array() } }],
      'contains binary data outside its source',
    );
  });

  it('requires both exact source and data property names for binary data', () => {
    expectInvalid(
      { notSource: { data: new Uint8Array() } },
      'contains binary data outside its source',
    );
    expectInvalid(
      { source: { notData: new Uint8Array() } },
      'contains binary data outside its source',
    );
  });
});

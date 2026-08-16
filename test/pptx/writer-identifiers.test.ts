import { describe, expect, it } from 'vitest';

import { createFieldIdAllocator } from '../../src/formats/pptx/writer/identifiers';

describe('PowerPoint writer identifiers', () => {
  it('allocates deterministic GUID-shaped field identities', () => {
    const allocator = createFieldIdAllocator();

    expect(allocator.allocateFieldId()).toBe(
      '{00000000-0000-0000-0000-000000000001}',
    );
    expect(allocator.allocateFieldId()).toBe(
      '{00000000-0000-0000-0000-000000000002}',
    );
  });

  it('uses uppercase hexadecimal suffixes with fixed width', () => {
    const allocator = createFieldIdAllocator(10);

    expect(allocator.allocateFieldId()).toBe(
      '{00000000-0000-0000-0000-00000000000A}',
    );
    expect(allocator.allocateFieldId()).toBe(
      '{00000000-0000-0000-0000-00000000000B}',
    );
  });

  it('keeps concurrent writer contexts isolated', () => {
    const first = createFieldIdAllocator();
    const second = createFieldIdAllocator();

    expect(first.allocateFieldId()).toBe(
      '{00000000-0000-0000-0000-000000000001}',
    );
    expect(first.allocateFieldId()).toBe(
      '{00000000-0000-0000-0000-000000000002}',
    );
    expect(second.allocateFieldId()).toBe(
      '{00000000-0000-0000-0000-000000000001}',
    );
    expect(first.allocateFieldId()).toBe(
      '{00000000-0000-0000-0000-000000000003}',
    );
    expect(second.allocateFieldId()).toBe(
      '{00000000-0000-0000-0000-000000000002}',
    );
  });

  it.each([
    0,
    -1,
    1.5,
    Number.NaN,
    Number.POSITIVE_INFINITY,
    0x1_0000_0000_0000,
  ])('rejects invalid starting identity %s', (start) => {
    expect(() => createFieldIdAllocator(start)).toThrow(
      new RangeError(
        'PowerPoint field identity start must be an integer from 1 through 281474976710655',
      ),
    );
  });

  it('allocates the maximum suffix exactly once', () => {
    const allocator = createFieldIdAllocator(0xffff_ffff_ffff);

    expect(allocator.allocateFieldId()).toBe(
      '{00000000-0000-0000-0000-FFFFFFFFFFFF}',
    );
    expect(() => allocator.allocateFieldId()).toThrow(
      new RangeError('PowerPoint field identity space is exhausted'),
    );
  });
});

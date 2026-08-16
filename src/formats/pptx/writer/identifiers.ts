import type { PptxTextSerializationContext } from './text-node';

const MAX_GUID_SUFFIX = 0xffff_ffff_ffff;

function assertStart(start: number): void {
  if (!Number.isSafeInteger(start) || start < 1 || start > MAX_GUID_SUFFIX) {
    throw new RangeError(
      'PowerPoint field identity start must be an integer from 1 through 281474976710655',
    );
  }
}

export function createFieldIdAllocator(
  start = 1,
): PptxTextSerializationContext {
  assertStart(start);
  let next = start;
  return {
    allocateFieldId(): string {
      if (next > MAX_GUID_SUFFIX) {
        throw new RangeError('PowerPoint field identity space is exhausted');
      }
      const suffix = next.toString(16).toUpperCase().padStart(12, '0');
      next += 1;
      return `{00000000-0000-0000-0000-${suffix}}`;
    },
  };
}

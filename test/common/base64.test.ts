import { describe, expect, it } from 'vitest';

import {
  base64ArrayBuffer,
  encodeBase64,
} from '../../src/common/binary/base64';

describe('runtime-neutral base64 encoding', () => {
  it.each([
    [[], ''],
    [[0], 'AA=='],
    [[255], '/w=='],
    [[0, 1], 'AAE='],
    [[255, 254], '//4='],
    [[0, 1, 2], 'AAEC'],
    [[72, 101, 108, 108, 111], 'SGVsbG8='],
    [[0, 1, 2, 253, 254, 255], 'AAEC/f7/'],
  ])('encodes %j exactly', (values, expected) => {
    expect(encodeBase64(Uint8Array.from(values))).toBe(expected);
  });

  it('encodes only the selected ArrayBufferView range', () => {
    const backing = Uint8Array.from([99, 72, 101, 108, 108, 111, 88]);
    const view = new DataView(backing.buffer, 1, 5);

    expect(encodeBase64(view)).toBe('SGVsbG8=');
  });

  it('accepts an ArrayBuffer without mutating caller-owned bytes', () => {
    const bytes = Uint8Array.from([0, 127, 128, 255]);
    const before = bytes.slice();

    expect(encodeBase64(bytes.buffer)).toBe('AH+A/w==');
    expect(bytes).toEqual(before);
  });

  it('keeps the compatibility export behavior identical', () => {
    const bytes = Uint8Array.from([1, 2, 3, 4]);
    expect(base64ArrayBuffer(bytes)).toBe(encodeBase64(bytes));
    expect(base64ArrayBuffer).toBe(encodeBase64);
  });
});

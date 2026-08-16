import { describe, expect, it } from 'vitest';

import {
  base64ArrayBuffer,
  decodeBase64,
  decodedBase64ByteLength,
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

  it.each([
    ['', []],
    ['AA==', [0]],
    ['/w==', [255]],
    ['AAE=', [0, 1]],
    ['//4=', [255, 254]],
    ['AAEC', [0, 1, 2]],
    ['SGVsbG8=', [72, 101, 108, 108, 111]],
    ['AAEC/f7/', [0, 1, 2, 253, 254, 255]],
  ])('decodes canonical %s exactly', (encoded, expected) => {
    expect(decodedBase64ByteLength(encoded)).toBe(expected.length);
    expect(Array.from(decodeBase64(encoded))).toEqual(expected);
  });

  it.each([
    'A',
    'AA',
    'AAA',
    ' AAE=',
    'AAE= ',
    'AA\nE=',
    'AA-E',
    'AA_E',
    '{AAA',
    '=AAA',
    'A=AA',
    'AA=A',
    'AA===',
    'AAAA=',
    'AAAAx',
    'xAAAA',
  ])('rejects malformed Base64 %j', (encoded) => {
    expect(() => decodedBase64ByteLength(encoded)).toThrow(
      'Invalid canonical Base64 encoding',
    );
    expect(() => decodeBase64(encoded)).toThrow(
      'Invalid canonical Base64 encoding',
    );
  });

  it.each(['AB==', 'A/==', 'AAF=', 'AA/=', '//9='])(
    'rejects non-canonical padding bits in %s',
    (encoded) => {
      expect(() => decodedBase64ByteLength(encoded)).toThrow(
        'Invalid canonical Base64 encoding',
      );
    },
  );

  it('round-trips every byte value without sharing output storage', () => {
    const input = Uint8Array.from({ length: 256 }, (_value, index) => index);
    const decoded = decodeBase64(encodeBase64(input));

    expect(decoded).toEqual(input);
    expect(decoded).not.toBe(input);
    decoded.fill(0);
    expect(input[255]).toBe(255);
  });

  it('accepts every RFC 4648 alphabet character', () => {
    const encoded =
      'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

    expect(decodedBase64ByteLength(encoded)).toBe(48);
    expect(encodeBase64(decodeBase64(encoded))).toBe(encoded);
  });

  it('validates a large canonical value without recursive regular expressions', () => {
    const encoded = 'A'.repeat(4_500_000);

    expect(decodedBase64ByteLength(encoded)).toBe(3_375_000);
  }, 15_000);
});

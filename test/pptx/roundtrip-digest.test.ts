import { createHash } from 'node:crypto';

import { describe, expect, it } from 'vitest';

import {
  canonicalSha256,
  sha256Bytes,
  sha256Text,
} from '../../src/formats/pptx/roundtrip/digest';

function nodeSha256(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

describe('PowerPoint round-trip SHA-256', () => {
  it.each([
    new Uint8Array(),
    new Uint8Array([0, 1, 2, 127, 128, 255]),
    new TextEncoder().encode('PowerPoint round-trip 🔒'),
  ])('matches an independent SHA-256 implementation', async (bytes) => {
    expect(await sha256Bytes(bytes)).toBe(nodeSha256(bytes));
  });

  it('hashes UTF-8 text rather than UTF-16 storage', async () => {
    const value = 'Việt Nam 🌏';
    expect(await sha256Text(value)).toBe(
      nodeSha256(new TextEncoder().encode(value)),
    );
  });

  it('hashes canonical object order identically', async () => {
    expect(await canonicalSha256({ b: 2, a: 1 })).toBe(
      await canonicalSha256({ a: 1, b: 2 }),
    );
    expect(await canonicalSha256({ a: 1, b: 2 })).toBe(
      nodeSha256(new TextEncoder().encode('{"a":1,"b":2}')),
    );
  });
});

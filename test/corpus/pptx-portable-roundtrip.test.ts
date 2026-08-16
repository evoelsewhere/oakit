import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';

import { describe, expect, it } from 'vitest';

import {
  parsePptx,
  parsePptxRoundTripJson,
  readPptxRoundTrip,
  serializePptxRoundTripJson,
  writePptxRoundTrip,
} from '../../src';

interface ResolvedCorpusEntry {
  id: string;
  path: string;
  producer: string;
  tier: 'curated' | 'large';
}

interface ResolvedCorpus {
  entries: ResolvedCorpusEntry[];
  version: number;
}

function sha256(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.byteLength !== right.byteLength) return false;
  for (let index = 0; index < left.byteLength; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}

const resolvedManifest = JSON.parse(
  await readFile(resolve('.cache/pptx-corpus/resolved.json'), 'utf8'),
) as ResolvedCorpus;

describe(`portable round-trip producer corpus v${resolvedManifest.version}`, () => {
  for (const entry of resolvedManifest.entries) {
    it(`${entry.producer}: ${entry.id}`, async () => {
      const source = new Uint8Array(await readFile(entry.path));
      const sourceDigest = sha256(source);
      const baseline = await parsePptx(source, {
        audioMode: 'none',
        errorMode: 'strict',
        imageMode: 'none',
        limits: { maxInputBytes: 150 * 1024 * 1024 },
        videoMode: 'none',
      });
      const runtime = await readPptxRoundTrip(source, {
        limits: { maxInputBytes: 150 * 1024 * 1024 },
      });
      const portable = await serializePptxRoundTripJson(runtime, {
        maxBase64Characters: 200 * 1024 * 1024,
        maxDecodedBytes: 150 * 1024 * 1024,
      });
      const wireValue: unknown = JSON.parse(JSON.stringify(portable));
      const restored = await parsePptxRoundTripJson(wireValue, {
        maxBase64Characters: 200 * 1024 * 1024,
        maxDecodedBytes: 150 * 1024 * 1024,
      });
      const output = await writePptxRoundTrip(restored, {
        limits: { maxInputBytes: 150 * 1024 * 1024 },
      });

      expect(runtime.source.sha256).toBe(sourceDigest);
      expect(portable.source.sha256).toBe(sourceDigest);
      expect(output.data.byteLength).toBe(source.byteLength);
      expect(bytesEqual(output.data, source)).toBe(true);
      expect(sha256(output.data)).toBe(sourceDigest);
      expect(output.report.level).toBe('R0');
      const reparsed = await parsePptx(output.data, {
        audioMode: 'none',
        errorMode: 'strict',
        imageMode: 'none',
        limits: { maxInputBytes: 150 * 1024 * 1024 },
        videoMode: 'none',
      });
      expect(reparsed).toEqual(baseline);
    });
  }
});

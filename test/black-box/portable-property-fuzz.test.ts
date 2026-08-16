import fc from 'fast-check';
import { beforeAll, describe, expect, it } from 'vitest';

import {
  parsePptxRoundTripJson,
  PptxWriteError,
  readPptxRoundTrip,
  serializePptxRoundTripJson,
  writePptxRoundTrip,
  type PptxRoundTripPortableJson,
} from '../../src';
import { createIndependentPptx } from './pptx-package';

const PORTABLE_FUZZ_SEED = 0x50_50_54;
const PORTABLE_FUZZ_RUNS = 64;
const BASE64_ALPHABET =
  'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
const ROOT_KEYS: string[] = [
  'consistency',
  'document',
  'format',
  'operations',
  'schemaVersion',
  'source',
  'supportProfile',
];
const SOURCE_KEYS: string[] = [
  'byteLength',
  'conformance',
  'kind',
  'packageBase64',
  'sha256',
];

let sourceBytes: Uint8Array;
let baseline: PptxRoundTripPortableJson;

function cloneBaseline(): PptxRoundTripPortableJson {
  return structuredClone(baseline);
}

function reorder(
  value: object,
  keys: readonly string[],
): Record<string, unknown> {
  const record = value as unknown as Readonly<Record<string, unknown>>;
  return Object.fromEntries(keys.map((key) => [key, record[key]]));
}

beforeAll(async () => {
  sourceBytes = await createIndependentPptx({
    'customXml/fuzz-sentinel.xml':
      '<?xml version="1.0"?><sentinel xmlns="urn:oakit:fuzz">preserve</sentinel>',
  });
  baseline = await serializePptxRoundTripJson(
    await readPptxRoundTrip(sourceBytes),
  );
});

describe('PowerPoint portable JSON seeded properties', () => {
  it('rejects arbitrary JSON values that do not satisfy the envelope contract', async () => {
    await fc.assert(
      fc.asyncProperty(fc.jsonValue(), async (value) => {
        let received: unknown;
        try {
          await parsePptxRoundTripJson(value);
        } catch (error) {
          received = error;
        }
        expect(received).toBeInstanceOf(PptxWriteError);
        expect(received).toMatchObject({ code: 'invalid-snapshot' });
      }),
      { numRuns: PORTABLE_FUZZ_RUNS, seed: PORTABLE_FUZZ_SEED },
    );
  });

  it('rejects every generated canonical Base64 byte corruption', async () => {
    const contentLength = baseline.source.packageBase64.indexOf('=');
    const stableContentLength =
      contentLength === -1
        ? baseline.source.packageBase64.length
        : contentLength;

    await fc.assert(
      fc.asyncProperty(
        fc.integer({ min: 0, max: stableContentLength - 4 }),
        fc.integer({ min: 1, max: BASE64_ALPHABET.length - 1 }),
        async (index, offset) => {
          const portable = cloneBaseline();
          const current = portable.source.packageBase64.charAt(index);
          const replacement = BASE64_ALPHABET.charAt(
            (BASE64_ALPHABET.indexOf(current) + offset) %
              BASE64_ALPHABET.length,
          );
          portable.source.packageBase64 = `${portable.source.packageBase64.slice(0, index)}${replacement}${portable.source.packageBase64.slice(index + 1)}`;

          await expect(parsePptxRoundTripJson(portable)).rejects.toMatchObject({
            code: 'snapshot-consistency-failed',
            message:
              'PowerPoint round-trip source SHA-256 does not match the snapshot',
          });
        },
      ),
      { numRuns: PORTABLE_FUZZ_RUNS, seed: PORTABLE_FUZZ_SEED + 1 },
    );
  });

  it('rejects every generated Base64 truncation without returning bytes', async () => {
    await fc.assert(
      fc.asyncProperty(fc.integer({ min: 1, max: 64 }), async (amount) => {
        const portable = cloneBaseline();
        portable.source.packageBase64 = portable.source.packageBase64.slice(
          0,
          -amount,
        );

        await expect(parsePptxRoundTripJson(portable)).rejects.toMatchObject({
          code: 'invalid-snapshot',
        });
      }),
      { numRuns: PORTABLE_FUZZ_RUNS, seed: PORTABLE_FUZZ_SEED + 2 },
    );
  });

  it('accepts every generated root and source key ordering after JSON transport', async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.shuffledSubarray(ROOT_KEYS, {
          maxLength: ROOT_KEYS.length,
          minLength: ROOT_KEYS.length,
        }),
        fc.shuffledSubarray(SOURCE_KEYS, {
          maxLength: SOURCE_KEYS.length,
          minLength: SOURCE_KEYS.length,
        }),
        async (rootKeys, sourceKeys) => {
          const portable = cloneBaseline();
          const source = reorder(portable.source, sourceKeys);
          const root = reorder(portable, rootKeys);
          root.source = source;
          const wireValue: unknown = JSON.parse(JSON.stringify(root));

          const restored = await parsePptxRoundTripJson(wireValue);
          const output = await writePptxRoundTrip(restored);

          expect(output.data).toEqual(sourceBytes);
        },
      ),
      { numRuns: PORTABLE_FUZZ_RUNS, seed: PORTABLE_FUZZ_SEED + 3 },
    );
  });
});

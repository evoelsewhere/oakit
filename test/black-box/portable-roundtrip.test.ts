import { Buffer } from 'node:buffer';

import { describe, expect, it } from 'vitest';

import {
  createPptx,
  defaultPptxRoundTripPortableLimits,
  parsePptxRoundTripJson,
  PptxRoundTripPortableLimitError,
  readPptxRoundTrip,
  serializePptxRoundTripJson,
  writePptxRoundTrip,
  type PptxRoundTripPortableJson,
  type PptxRoundTripSnapshot,
  type PptxSceneDocument,
} from '../../src';

function scene(): PptxSceneDocument {
  return {
    layouts: [],
    masters: [],
    media: [],
    schemaVersion: 2,
    size: { height: 540, width: 960 },
    slides: [
      {
        elements: [
          {
            authored: {
              transform: { height: 80, width: 300, x: 20, y: 30 },
            },
            key: 'agent-text',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    {
                      key: 'agent-run',
                      text: 'Portable agent snapshot',
                      type: 'run',
                    },
                  ],
                  key: 'agent-paragraph',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'agent-slide',
      },
    ],
    themes: [],
  };
}

async function runtimeFixture(): Promise<{
  bytes: Uint8Array;
  snapshot: PptxRoundTripSnapshot;
}> {
  const created = await createPptx(scene());
  return {
    bytes: created.data,
    snapshot: await readPptxRoundTrip(created.data),
  };
}

async function portableFixture(): Promise<{
  bytes: Uint8Array;
  portable: PptxRoundTripPortableJson;
}> {
  const { bytes, snapshot } = await runtimeFixture();
  return {
    bytes,
    portable: await serializePptxRoundTripJson(snapshot),
  };
}

async function expectPortableError(
  value: unknown,
  code: string,
  messagePart: string,
): Promise<void> {
  let received: unknown;
  try {
    await parsePptxRoundTripJson(value);
  } catch (error) {
    received = error;
  }
  expect(received).toMatchObject({ code });
  if (!(received instanceof Error)) {
    throw new Error('Expected a portable snapshot error');
  }
  expect(received.message).toContain(messagePart);
}

function record(value: unknown): Record<string, unknown> {
  return value as Record<string, unknown>;
}

function portableSource(
  value: PptxRoundTripPortableJson,
): Record<string, unknown> {
  return value.source as unknown as Record<string, unknown>;
}

describe('PowerPoint portable JSON round trips through the public API', () => {
  it('preserves exact source bytes through ordinary JSON stringify and parse', async () => {
    const { bytes, snapshot } = await runtimeFixture();

    const portable = await serializePptxRoundTripJson(snapshot);
    const json = JSON.stringify(portable);
    const plainValue: unknown = JSON.parse(json);
    const restored = await parsePptxRoundTripJson(plainValue);
    const result = await writePptxRoundTrip(restored);

    expect(Object.keys(portable)).toEqual([
      'consistency',
      'document',
      'format',
      'operations',
      'schemaVersion',
      'source',
      'supportProfile',
    ]);
    expect(Object.keys(portable.source)).toEqual([
      'byteLength',
      'conformance',
      'kind',
      'packageBase64',
      'sha256',
    ]);
    expect(portable.source).toMatchObject({
      byteLength: bytes.byteLength,
      conformance: 'transitional',
      kind: 'base64',
      packageBase64: Buffer.from(bytes).toString('base64'),
      sha256: snapshot.source.sha256,
    });
    expect(JSON.parse(json)).toEqual(portable);
    expect(restored.source.data).toBeInstanceOf(Uint8Array);
    expect(restored.source.data).toEqual(bytes);
    expect(result.data).toEqual(bytes);
    expect(result.report.level).toBe('R0');
  });

  it('serializes Blob source without leaking a runtime-only value into JSON', async () => {
    const created = await createPptx(scene());
    const blobBytes = new Uint8Array(created.data);
    const snapshot = await readPptxRoundTrip(
      new Blob([blobBytes.buffer], {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      }),
    );

    const portable = await serializePptxRoundTripJson(snapshot);

    expect(snapshot.source.data).toBeInstanceOf(Blob);
    expect(JSON.stringify(portable)).not.toContain('Blob');
    expect(portable.source.packageBase64).toBe(
      Buffer.from(created.data).toString('base64'),
    );
  });

  it('owns runtime input synchronously before caller mutation', async () => {
    const { bytes, snapshot } = await runtimeFixture();
    if (!(snapshot.source.data instanceof Uint8Array)) {
      throw new Error('Expected runtime bytes');
    }

    const serializing = serializePptxRoundTripJson(snapshot);
    snapshot.source.data.fill(0);
    snapshot.document.size.width = 1;
    snapshot.consistency.operationsSha256 = 'b'.repeat(64);

    const portable = await serializing;
    expect(portable.source.packageBase64).toBe(
      Buffer.from(bytes).toString('base64'),
    );
    expect(portable.document.size.width).toBe(960);
  });

  it('owns portable input synchronously before caller mutation', async () => {
    const { bytes, portable } = await portableFixture();

    const parsing = parsePptxRoundTripJson(portable);
    portable.source.packageBase64 = 'AAAA';
    portable.document.size.width = 1;
    portable.consistency.operationsSha256 = 'b'.repeat(64);

    const snapshot = await parsing;
    expect(snapshot.source.data).toEqual(bytes);
    expect(snapshot.document.size.width).toBe(960);
  });

  it('is deterministic and isolated across concurrent portable conversions', async () => {
    const { bytes, snapshot } = await runtimeFixture();
    const portableValues = await Promise.all([
      serializePptxRoundTripJson(snapshot),
      serializePptxRoundTripJson(snapshot),
      serializePptxRoundTripJson(snapshot),
    ]);
    expect(portableValues[1]).toEqual(portableValues[0]);
    expect(portableValues[2]).toEqual(portableValues[0]);

    const restored = await Promise.all(
      portableValues.map((portable) => parsePptxRoundTripJson(portable)),
    );
    expect(restored[0]?.source.data).toEqual(bytes);
    expect(restored[1]?.source.data).toEqual(bytes);
    expect(restored[2]?.source.data).toEqual(bytes);
    if (!(restored[0]?.source.data instanceof Uint8Array)) {
      throw new Error('Expected restored bytes');
    }
    restored[0].source.data.fill(0);
    expect(restored[1]?.source.data).toEqual(bytes);
    expect(restored[2]?.source.data).toEqual(bytes);
  });

  it('publishes explicit portable defaults', () => {
    expect(defaultPptxRoundTripPortableLimits()).toEqual({
      maxBase64Characters: 139_810_136,
      maxDecodedBytes: 104_857_600,
    });
  });

  it('accepts exact portable limits and rejects each value one below', async () => {
    const { bytes, snapshot } = await runtimeFixture();
    const encodedLength = Buffer.from(bytes).toString('base64').length;

    const portable = await serializePptxRoundTripJson(snapshot, {
      maxBase64Characters: encodedLength,
      maxDecodedBytes: bytes.byteLength,
    });
    await expect(
      parsePptxRoundTripJson(portable, {
        maxBase64Characters: encodedLength,
        maxDecodedBytes: bytes.byteLength,
      }),
    ).resolves.toMatchObject({ source: { byteLength: bytes.byteLength } });

    for (const [limitName, actual, limit] of [
      ['maxBase64Characters', encodedLength, encodedLength - 1],
      ['maxDecodedBytes', bytes.byteLength, bytes.byteLength - 1],
    ] as const) {
      let received: unknown;
      try {
        await serializePptxRoundTripJson(snapshot, { [limitName]: limit });
      } catch (error) {
        received = error;
      }
      expect(received).toBeInstanceOf(PptxRoundTripPortableLimitError);
      expect(received).toMatchObject({
        actual,
        limit,
        limitName,
        message: `PowerPoint portable snapshot limit ${limitName} exceeded: ${actual} > ${limit}`,
        name: 'PptxRoundTripPortableLimitError',
      });
    }
  });

  it.each([0, -1, 1.5, Number.MAX_SAFE_INTEGER + 1])(
    'rejects an invalid portable limit %s',
    async (limit) => {
      const { snapshot } = await runtimeFixture();
      await expect(
        serializePptxRoundTripJson(snapshot, {
          maxBase64Characters: limit,
        }),
      ).rejects.toThrow(
        'PowerPoint portable snapshot limit maxBase64Characters must be a positive integer',
      );
      await expect(
        serializePptxRoundTripJson(snapshot, { maxDecodedBytes: limit }),
      ).rejects.toThrow(
        'PowerPoint portable snapshot limit maxDecodedBytes must be a positive integer',
      );
    },
  );

  it.each([
    [
      'primitive root',
      (portable: PptxRoundTripPortableJson) => {
        void portable;
        return null;
      },
      'invalid root shape',
    ],
    [
      'extra root field',
      (portable: PptxRoundTripPortableJson) => {
        record(portable).extra = true;
        return portable;
      },
      'invalid root shape',
    ],
    [
      'substituted root field',
      (portable: PptxRoundTripPortableJson) => {
        delete record(portable).document;
        record(portable).extra = true;
        return portable;
      },
      'invalid root shape',
    ],
    [
      'source shape',
      (portable: PptxRoundTripPortableJson) => {
        record(portable).source = {};
        return portable;
      },
      'source has an invalid shape',
    ],
    [
      'substituted source field',
      (portable: PptxRoundTripPortableJson) => {
        delete portableSource(portable).sha256;
        portableSource(portable).extra = true;
        return portable;
      },
      'source has an invalid shape',
    ],
    [
      'source kind',
      (portable: PptxRoundTripPortableJson) => {
        portableSource(portable).kind = 'bytes';
        return portable;
      },
      'source kind must be base64',
    ],
    [
      'source data type',
      (portable: PptxRoundTripPortableJson) => {
        portableSource(portable).packageBase64 = 1;
        return portable;
      },
      'packageBase64 must be a string',
    ],
  ])('rejects portable %s', async (_name, create, message) => {
    const value = create((await portableFixture()).portable);
    await expectPortableError(value, 'invalid-snapshot', message);
  });

  it.each(['AA A=', 'AAAAx', 'xAAAA', 'AB==', 'AAF='])(
    'rejects non-canonical package Base64 %j with its cause',
    async (packageBase64) => {
      const { portable } = await portableFixture();
      portable.source.packageBase64 = packageBase64;

      let received: unknown;
      try {
        await parsePptxRoundTripJson(portable);
      } catch (error) {
        received = error;
      }
      expect(received).toMatchObject({
        code: 'invalid-snapshot',
        message:
          'PowerPoint portable snapshot packageBase64 must be canonical Base64',
      });
      if (!(received instanceof Error)) {
        throw new Error('Expected a portable snapshot error');
      }
      expect(received.cause).toBeInstanceOf(RangeError);
    },
  );

  it.each([
    ['zero', 0],
    ['negative', -1],
    ['fractional', 1.5],
    ['unsafe', Number.MAX_SAFE_INTEGER + 1],
  ])('rejects a %s declared byte length', async (_name, byteLength) => {
    const { portable } = await portableFixture();
    portableSource(portable).byteLength = byteLength;

    await expect(parsePptxRoundTripJson(portable)).rejects.toMatchObject({
      code: 'invalid-snapshot',
      message:
        'PowerPoint portable snapshot source byteLength must be a positive safe integer',
    });
  });

  it('rejects a positive declared length that differs from decoded bytes', async () => {
    const { portable } = await portableFixture();
    portable.source.byteLength += 1;

    await expect(parsePptxRoundTripJson(portable)).rejects.toMatchObject({
      code: 'invalid-snapshot',
      message:
        'PowerPoint portable snapshot source byteLength does not match packageBase64',
    });
  });

  it('rejects a portable string limit before Base64 validation', async () => {
    const { portable } = await portableFixture();
    portable.source.packageBase64 = 'not canonical base64';

    await expect(
      parsePptxRoundTripJson(portable, { maxBase64Characters: 1 }),
    ).rejects.toMatchObject({
      actual: portable.source.packageBase64.length,
      limit: 1,
      limitName: 'maxBase64Characters',
    });
  });

  it('rejects decoded byte limits before allocating runtime state', async () => {
    const { bytes, portable } = await portableFixture();

    await expect(
      parsePptxRoundTripJson(portable, {
        maxDecodedBytes: bytes.byteLength - 1,
      }),
    ).rejects.toMatchObject({
      actual: bytes.byteLength,
      limit: bytes.byteLength - 1,
      limitName: 'maxDecodedBytes',
    });
  });

  it.each([
    [
      'source digest',
      (portable: PptxRoundTripPortableJson) => {
        portable.source.sha256 = 'b'.repeat(64);
      },
      'PowerPoint round-trip source SHA-256 does not match the snapshot',
    ],
    [
      'preview',
      (portable: PptxRoundTripPortableJson) => {
        portable.document.size.width += 1;
      },
      'PowerPoint round-trip snapshot consistency does not match its bound state',
    ],
    [
      'consistency digest',
      (portable: PptxRoundTripPortableJson) => {
        portable.consistency.semanticPreviewSha256 = 'b'.repeat(64);
      },
      'PowerPoint round-trip snapshot consistency does not match its bound state',
    ],
  ])('rejects modified portable %s', async (_name, mutate, message) => {
    const { portable } = await portableFixture();
    mutate(portable);

    await expect(parsePptxRoundTripJson(portable)).rejects.toMatchObject({
      code: 'snapshot-consistency-failed',
      message,
    });
  });
});

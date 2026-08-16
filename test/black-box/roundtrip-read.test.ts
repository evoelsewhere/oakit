import { createHash } from 'node:crypto';

import { describe, expect, it } from 'vitest';

import {
  createPptx,
  PptxParseError,
  readPptxRoundTrip,
  type PptxRoundTripReadOptions,
  type PptxSceneDocument,
} from '../../src';
import { createMinimalPptx } from '../pptx/fixture';

function sha256(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

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
            key: 'source-text',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [
                    { key: 'source-run', text: 'Snapshot', type: 'run' },
                  ],
                  key: 'source-paragraph',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'source-slide',
      },
    ],
    themes: [],
  };
}

describe('PowerPoint round-trip reading through the public API', () => {
  it('constructs a sealed R0 snapshot from an owned byte copy', async () => {
    const created = await createPptx(scene());
    const snapshot = await readPptxRoundTrip(created.data);

    expect(snapshot).toMatchObject({
      format: 'pptx',
      operations: [],
      schemaVersion: 1,
      source: {
        byteLength: created.data.byteLength,
        conformance: 'transitional',
        kind: 'bytes',
        sha256: sha256(created.data),
      },
      supportProfile: {
        effectiveLevel: 'R0',
        id: 'pptx-roundtrip-r0',
        producerMatrix: [],
        version: '1',
      },
    });
    expect(snapshot.source.data).toBeInstanceOf(Uint8Array);
    expect(snapshot.source.data).not.toBe(created.data);
    expect(snapshot.source.data).toEqual(created.data);
    expect(snapshot.document.size).toEqual(scene().size);
    expect(snapshot.document.slides).toHaveLength(1);
    expect(snapshot.document.slides[0]?.elements[0]).toMatchObject({
      key: 'slide-1-element-1',
      text: {
        paragraphs: [
          {
            children: [
              {
                key: 'slide-1-element-1-run-1',
                text: 'Snapshot',
                type: 'run',
              },
            ],
          },
        ],
      },
      type: 'text',
    });
    expect(snapshot.consistency).toMatchObject({
      canonicalizationVersion: 'canonical-json-v1',
      capabilityProfileVersion: 'pptx-roundtrip-r0-v1',
      contractVersion: '1',
      hashAlgorithm: 'sha256',
      keyAlgorithmVersion: 'pptx-scene-key-v1',
    });
    expect(Object.values(snapshot.consistency)).not.toContain('');
  });

  it('copies a byte view before concurrent caller mutation', async () => {
    const created = await createPptx(scene());
    const owner = new Uint8Array(created.data.byteLength + 2);
    owner.set(created.data, 1);
    const view = owner.subarray(1, owner.byteLength - 1);
    const reading = readPptxRoundTrip(view);
    owner.fill(0);

    const snapshot = await reading;
    expect(snapshot.source.data).toEqual(created.data);
    expect(snapshot.source.sha256).toBe(sha256(created.data));
  });

  it('preserves Blob as the runtime transport', async () => {
    const created = await createPptx(scene());
    const blobBytes = new Uint8Array(created.data.byteLength);
    blobBytes.set(created.data);
    const blob = new Blob([blobBytes.buffer], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });
    const snapshot = await readPptxRoundTrip(blob);

    expect(snapshot.source.data).toBeInstanceOf(Blob);
    if (!(snapshot.source.data instanceof Blob)) {
      throw new Error('Expected snapshot Blob');
    }
    expect(snapshot.source.data).not.toBe(blob);
    expect(snapshot.source.data.type).toBe(blob.type);
    expect(new Uint8Array(await snapshot.source.data.arrayBuffer())).toEqual(
      created.data,
    );
  });

  it('cannot be weakened to tolerant parsing by an extra runtime option', async () => {
    const malformed = await createMinimalPptx({
      'ppt/theme/theme1.xml': '<a:theme>',
    });
    const untrustedOptions = {
      errorMode: 'tolerant',
    } as unknown as PptxRoundTripReadOptions;

    await expect(
      readPptxRoundTrip(malformed, untrustedOptions),
    ).rejects.toBeInstanceOf(PptxParseError);
  });

  it('applies compressed input limits before snapshot construction', async () => {
    const created = await createPptx(scene());

    await expect(
      readPptxRoundTrip(created.data, {
        limits: { maxInputBytes: created.data.byteLength - 1 },
      }),
    ).rejects.toMatchObject({
      actual: created.data.byteLength,
      limit: created.data.byteLength - 1,
      limitName: 'maxInputBytes',
    });
  });
});

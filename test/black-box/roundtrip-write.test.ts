import { createHash } from 'node:crypto';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import {
  createPptx,
  parsePptx,
  readPptxRoundTrip,
  writePptxRoundTrip,
  type PptxRoundTripSnapshot,
  type PptxSceneDocument,
} from '../../src';

function sha256(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

function scene(text = 'Exact source'): PptxSceneDocument {
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
            key: 'text',
            resolved: { hidden: false },
            text: {
              body: {},
              paragraphs: [
                {
                  children: [{ key: 'run', text, type: 'run' }],
                  key: 'paragraph',
                },
              ],
            },
            type: 'text',
          },
        ],
        key: 'slide',
      },
    ],
    themes: [],
  };
}

async function sourceSnapshot(): Promise<{
  bytes: Uint8Array;
  snapshot: PptxRoundTripSnapshot;
}> {
  const created = await createPptx(scene());
  return {
    bytes: created.data,
    snapshot: await readPptxRoundTrip(created.data),
  };
}

function record(value: unknown): Record<string, unknown> {
  return value as Record<string, unknown>;
}

describe('PowerPoint exact round-trip writing through the public API', () => {
  it('returns an independent byte-identical package with an exact R0 report', async () => {
    const { bytes, snapshot } = await sourceSnapshot();
    const archive = await JSZip.loadAsync(bytes);
    const partCount = Object.values(archive.files).filter(
      (part) => !part.dir,
    ).length;

    const result = await writePptxRoundTrip(snapshot);

    expect(result.data).not.toBe(bytes);
    expect(result.data).not.toBe(snapshot.source.data);
    expect(result.data).toEqual(bytes);
    expect(result.data.byteLength).toBe(bytes.byteLength);
    expect(sha256(result.data)).toBe(sha256(bytes));
    expect(result.report).toEqual({
      addedPartCount: 0,
      copiedPartCount: partCount,
      diagnostics: [],
      level: 'R0',
      operations: [],
      patchedPartCount: 0,
      producerEvidence: [],
      rebuiltPartCount: 0,
      removedPartCount: 0,
      supportProfile: {
        effectiveLevel: 'R0',
        id: 'pptx-roundtrip-r0',
        producerMatrix: [],
        version: '1',
      },
    });
    await expect(
      parsePptx(result.data, {
        audioMode: 'none',
        errorMode: 'strict',
        imageMode: 'none',
        videoMode: 'none',
      }),
    ).resolves.toMatchObject({ size: scene().size });
  });

  it('materializes Blob source without changing exact package bytes', async () => {
    const created = await createPptx(scene());
    const blobBytes = new Uint8Array(created.data);
    const blob = new Blob([blobBytes.buffer], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });
    const snapshot = await readPptxRoundTrip(blob);

    const result = await writePptxRoundTrip(snapshot);

    expect(snapshot.source.data).toBeInstanceOf(Blob);
    expect(result.data).toEqual(created.data);
    expect(sha256(result.data)).toBe(snapshot.source.sha256);
  });

  it('owns the complete snapshot synchronously before caller mutation', async () => {
    const { bytes, snapshot } = await sourceSnapshot();
    if (!(snapshot.source.data instanceof Uint8Array)) {
      throw new Error('Expected a byte snapshot');
    }

    const writing = writePptxRoundTrip(snapshot);
    snapshot.source.data.fill(0);
    snapshot.document.size.width = 1;
    snapshot.consistency.operationsSha256 = 'b'.repeat(64);

    const result = await writing;
    expect(result.data).toEqual(bytes);
    expect(result.report.level).toBe('R0');
  });

  it('is isolated and deterministic across concurrent calls', async () => {
    const { bytes, snapshot } = await sourceSnapshot();

    const [first, second, third] = await Promise.all([
      writePptxRoundTrip(snapshot),
      writePptxRoundTrip(snapshot),
      writePptxRoundTrip(snapshot),
    ]);

    expect(first.data).toEqual(bytes);
    expect(second.data).toEqual(first.data);
    expect(third.data).toEqual(first.data);
    expect(second.report).toEqual(first.report);
    expect(third.report).toEqual(first.report);
    first.data.fill(0);
    expect(second.data).toEqual(bytes);
    expect(third.data).toEqual(bytes);
  });

  it('rejects edit operations before reading source bytes', async () => {
    const { snapshot } = await sourceSnapshot();
    record(snapshot).operations = [{ type: 'replace-text' }];

    await expect(writePptxRoundTrip(snapshot)).rejects.toMatchObject({
      code: 'unsupported-edit-operation',
      message: 'PowerPoint R0 round-trip does not support edit operations',
    });
  });

  it('rejects mutated source bytes before parsing the package', async () => {
    const { snapshot } = await sourceSnapshot();
    if (!(snapshot.source.data instanceof Uint8Array)) {
      throw new Error('Expected a byte snapshot');
    }
    snapshot.source.data[0] = 0;

    await expect(writePptxRoundTrip(snapshot)).rejects.toMatchObject({
      code: 'snapshot-consistency-failed',
      message:
        'PowerPoint round-trip source SHA-256 does not match the snapshot',
    });
  });

  it('rejects a substituted lexical source digest', async () => {
    const { snapshot } = await sourceSnapshot();
    snapshot.source.sha256 = 'b'.repeat(64);

    await expect(writePptxRoundTrip(snapshot)).rejects.toMatchObject({
      code: 'snapshot-consistency-failed',
      message:
        'PowerPoint round-trip source SHA-256 does not match the snapshot',
    });
  });

  it.each([
    [
      'semantic preview',
      (snapshot: PptxRoundTripSnapshot) => {
        snapshot.document.size.width += 1;
      },
    ],
    [
      'source conformance',
      (snapshot: PptxRoundTripSnapshot) => {
        snapshot.source.conformance = 'strict';
      },
    ],
    [
      'operations digest',
      (snapshot: PptxRoundTripSnapshot) => {
        snapshot.consistency.operationsSha256 = 'b'.repeat(64);
      },
    ],
    [
      'preview digest',
      (snapshot: PptxRoundTripSnapshot) => {
        snapshot.consistency.semanticPreviewSha256 = 'b'.repeat(64);
      },
    ],
    [
      'source manifest digest',
      (snapshot: PptxRoundTripSnapshot) => {
        snapshot.consistency.sourceManifestSha256 = 'b'.repeat(64);
      },
    ],
  ])('rejects modified bound state: %s', async (_name, mutate) => {
    const { snapshot } = await sourceSnapshot();
    mutate(snapshot);

    await expect(writePptxRoundTrip(snapshot)).rejects.toMatchObject({
      code: 'snapshot-consistency-failed',
      message:
        'PowerPoint round-trip snapshot consistency does not match its bound state',
    });
  });

  it('enforces runtime input limits at the exact source boundary', async () => {
    const { bytes, snapshot } = await sourceSnapshot();

    await expect(
      writePptxRoundTrip(snapshot, {
        limits: { maxInputBytes: bytes.byteLength },
      }),
    ).resolves.toMatchObject({ data: bytes });
    await expect(
      writePptxRoundTrip(snapshot, {
        limits: { maxInputBytes: bytes.byteLength - 1 },
      }),
    ).rejects.toMatchObject({
      actual: bytes.byteLength,
      limit: bytes.byteLength - 1,
      limitName: 'maxInputBytes',
    });
  });
});

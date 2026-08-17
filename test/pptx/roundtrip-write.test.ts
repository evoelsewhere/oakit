import { createHash } from 'node:crypto';

import { describe, expect, it } from 'vitest';

import {
  createPptx,
  readPptxRoundTrip,
  replacePptxRoundTripText,
  writePptxRoundTrip,
  type PptxSceneDocument,
} from '../../src';
import { PptxParseError } from '../../src/formats/pptx/errors';
import { resolvePptxResourceLimits } from '../../src/formats/pptx/internal/resource-limits';
import { parse } from '../../src/formats/pptx/parser';
import { createPptxSnapshotConsistency } from '../../src/formats/pptx/roundtrip/consistency';
import { inspectPptxRoundTripPackage } from '../../src/formats/pptx/roundtrip/source';
import type { PptxRoundTripSnapshot } from '../../src/formats/pptx/roundtrip/types';
import { writePptxRoundTripWithDependencies } from '../../src/formats/pptx/roundtrip/write';
import { createMinimalPptx } from './fixture';

function sha256(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

async function bindSnapshot(snapshot: PptxRoundTripSnapshot): Promise<void> {
  snapshot.consistency = await createPptxSnapshotConsistency({
    document: snapshot.document,
    operations: snapshot.operations,
    source: snapshot.source,
    supportProfile: snapshot.supportProfile,
  });
}

function editableScene(): PptxSceneDocument {
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
                  children: [{ key: 'run', text: 'Before', type: 'run' }],
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

async function editedFixture() {
  const created = await createPptx(editableScene());
  const snapshot = await readPptxRoundTrip(created.data);
  return {
    bytes: created.data,
    snapshot: await replacePptxRoundTripText(snapshot, {
      targetKey: 'slide-1-element-1-run-1',
      value: 'After',
    }),
  };
}

describe('PowerPoint exact round-trip source verification', () => {
  it('rejects a strictly invalid source even when its claimed state is rebound', async () => {
    const snapshot = await readPptxRoundTrip(await createMinimalPptx());
    const malformed = await createMinimalPptx({
      'ppt/theme/theme1.xml': '<a:theme>',
    });
    snapshot.source.byteLength = malformed.byteLength;
    snapshot.source.conformance = 'unknown';
    snapshot.source.data = malformed;
    snapshot.source.sha256 = sha256(malformed);
    await bindSnapshot(snapshot);

    let received: unknown;
    try {
      await writePptxRoundTrip(snapshot);
    } catch (error) {
      received = error;
    }

    expect(received).toMatchObject({
      code: 'snapshot-consistency-failed',
      message: 'PowerPoint round-trip source failed strict verification',
    });
    if (!(received instanceof Error)) {
      throw new Error('Expected a verification error');
    }
    expect(received.cause).toBeInstanceOf(PptxParseError);
  });

  it('rejects a rebound semantic preview that no longer describes the source', async () => {
    const snapshot = await readPptxRoundTrip(await createMinimalPptx());
    snapshot.document.size.width += 1;
    await bindSnapshot(snapshot);

    await expect(writePptxRoundTrip(snapshot)).rejects.toMatchObject({
      code: 'snapshot-consistency-failed',
      message: 'PowerPoint round-trip source does not match the snapshot',
    });
  });

  it('rejects rebound conformance that disagrees with the source namespace', async () => {
    const snapshot = await readPptxRoundTrip(await createMinimalPptx());
    snapshot.source.conformance = 'strict';
    await bindSnapshot(snapshot);

    await expect(writePptxRoundTrip(snapshot)).rejects.toMatchObject({
      code: 'snapshot-consistency-failed',
      message: 'PowerPoint round-trip source does not match the snapshot',
    });
  });

  it('wraps an independently failing output parse with its exact options and cause', async () => {
    const { snapshot } = await editedFixture();
    const limits = resolvePptxResourceLimits();
    const cause = new Error('independent output parse failed');
    let parseCalls = 0;
    let outputOptions: Parameters<typeof parse>[1] | undefined;

    const writing = writePptxRoundTripWithDependencies(snapshot, limits, {
      inspect: inspectPptxRoundTripPackage,
      parse: async (data, options) => {
        parseCalls += 1;
        if (parseCalls === 2) {
          outputOptions = options;
          throw cause;
        }
        return parse(data, options);
      },
      patch: () =>
        Promise.resolve({
          copiedPartCount: 10,
          data: new Uint8Array([1, 2, 3]),
          patchedPartCount: 1,
        }),
    });

    await expect(writing).rejects.toMatchObject({
      cause,
      code: 'verification-failed',
      message: 'PowerPoint text edit output failed strict verification',
    });
    expect(outputOptions).toEqual({
      audioMode: 'none',
      errorMode: 'strict',
      imageMode: 'none',
      limits,
      videoMode: 'none',
    });
  });

  it('rejects a patched package whose strict preview misses the operation', async () => {
    const { bytes, snapshot } = await editedFixture();
    const limits = resolvePptxResourceLimits();

    await expect(
      writePptxRoundTripWithDependencies(snapshot, limits, {
        inspect: inspectPptxRoundTripPackage,
        parse,
        patch: () =>
          Promise.resolve({
            copiedPartCount: 10,
            data: bytes,
            patchedPartCount: 1,
          }),
      }),
    ).rejects.toMatchObject({
      code: 'verification-failed',
      message:
        'PowerPoint text edit output does not match the requested semantics',
    });
  });
});

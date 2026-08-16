import { createHash } from 'node:crypto';

import { describe, expect, it } from 'vitest';

import { readPptxRoundTrip, writePptxRoundTrip } from '../../src';
import { PptxParseError } from '../../src/formats/pptx/errors';
import { createPptxSnapshotConsistency } from '../../src/formats/pptx/roundtrip/consistency';
import type { PptxRoundTripSnapshot } from '../../src/formats/pptx/roundtrip/types';
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
});

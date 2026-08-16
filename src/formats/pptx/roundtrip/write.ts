import { parse } from '../parser';
import {
  resolvePptxResourceLimits,
  type ResolvedPptxResourceLimits,
} from '../internal/resource-limits';
import { PptxWriteError } from '../write-error';
import type { PptxWriteReport, PptxWriteResult } from '../write-types';
import {
  createPptxRoundTripSupportProfile,
  createPptxSnapshotConsistency,
} from './consistency';
import { createPowerPointRoundTripPreview } from './preview';
import {
  inspectPptxRoundTripPackage,
  normalizePptxRoundTripInput,
} from './source';
import type {
  PptxRoundTripSnapshot,
  PptxRoundTripWriteOptions,
  PptxSnapshotConsistency,
} from './types';
import { validatePptxRoundTripSnapshot } from './validate';

function consistencyMatches(
  actual: PptxSnapshotConsistency,
  expected: PptxSnapshotConsistency,
): boolean {
  return (
    actual.operationsSha256 === expected.operationsSha256 &&
    actual.semanticPreviewSha256 === expected.semanticPreviewSha256 &&
    actual.sourceManifestSha256 === expected.sourceManifestSha256
  );
}

function consistencyFailure(message: string, cause?: unknown): never {
  throw new PptxWriteError('snapshot-consistency-failed', message, { cause });
}

function r0Report(copiedPartCount: number): PptxWriteReport {
  return {
    addedPartCount: 0,
    copiedPartCount,
    diagnostics: [],
    level: 'R0',
    operations: [],
    patchedPartCount: 0,
    producerEvidence: [],
    rebuiltPartCount: 0,
    removedPartCount: 0,
    supportProfile: createPptxRoundTripSupportProfile(),
  };
}

export async function writePptxRoundTripWithLimits(
  value: PptxRoundTripSnapshot,
  limits: ResolvedPptxResourceLimits,
): Promise<PptxWriteResult> {
  const validated = validatePptxRoundTripSnapshot(value, limits);
  if (validated.operations.length !== 0) {
    throw new PptxWriteError(
      'unsupported-edit-operation',
      'PowerPoint text edit writing is not enabled',
    );
  }
  const snapshot = structuredClone(validated);
  const normalized = await normalizePptxRoundTripInput(
    snapshot.source.data,
    limits,
  );

  if (normalized.sha256 !== snapshot.source.sha256) {
    consistencyFailure(
      'PowerPoint round-trip source SHA-256 does not match the snapshot',
    );
  }

  const claimedConsistency = await createPptxSnapshotConsistency({
    document: snapshot.document,
    operations: snapshot.operations,
    source: snapshot.source,
    supportProfile: snapshot.supportProfile,
  });
  if (!consistencyMatches(snapshot.consistency, claimedConsistency)) {
    consistencyFailure(
      'PowerPoint round-trip snapshot consistency does not match its bound state',
    );
  }

  let parsed: Awaited<ReturnType<typeof parse>>;
  let inspection: Awaited<ReturnType<typeof inspectPptxRoundTripPackage>>;
  try {
    parsed = await parse(normalized.bytes, {
      audioMode: 'none',
      errorMode: 'strict',
      imageMode: 'none',
      limits,
      videoMode: 'none',
    });
    inspection = await inspectPptxRoundTripPackage(normalized.bytes, limits);
  } catch (cause) {
    consistencyFailure(
      'PowerPoint round-trip source failed strict verification',
      cause,
    );
  }

  const sourceDocument = createPowerPointRoundTripPreview(parsed);
  const sourceConsistency = await createPptxSnapshotConsistency({
    document: sourceDocument,
    operations: snapshot.operations,
    source: {
      byteLength: normalized.byteLength,
      conformance: inspection.conformance,
      sha256: normalized.sha256,
    },
    supportProfile: createPptxRoundTripSupportProfile(),
  });
  if (!consistencyMatches(snapshot.consistency, sourceConsistency)) {
    consistencyFailure(
      'PowerPoint round-trip source does not match the snapshot',
    );
  }

  return {
    data: normalized.bytes,
    report: r0Report(inspection.partCount),
  };
}

export async function writePptxRoundTrip(
  value: PptxRoundTripSnapshot,
  options: PptxRoundTripWriteOptions = {},
): Promise<PptxWriteResult> {
  return writePptxRoundTripWithLimits(
    value,
    resolvePptxResourceLimits(options.limits),
  );
}

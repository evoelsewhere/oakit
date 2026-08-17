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
import { canonicalJson } from './canonical-json';
import { applyPptxRoundTripOperationsToPreview } from './edit';
import { patchPptxTextOperations } from './patch-text';
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

function r2Report(
  snapshot: PptxRoundTripSnapshot,
  copiedPartCount: number,
  patchedPartCount: number,
): PptxWriteReport {
  return {
    addedPartCount: 0,
    copiedPartCount,
    diagnostics: [],
    level: 'R2',
    operations: snapshot.operations.map((operation) => ({
      id: operation.id,
      kind: operation.kind,
      status: 'verified',
    })),
    patchedPartCount,
    producerEvidence: [],
    rebuiltPartCount: 0,
    removedPartCount: 0,
    supportProfile: snapshot.supportProfile,
  };
}

export async function writePptxRoundTripWithLimits(
  value: PptxRoundTripSnapshot,
  limits: ResolvedPptxResourceLimits,
): Promise<PptxWriteResult> {
  const validated = validatePptxRoundTripSnapshot(value, limits);
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
    supportProfile: snapshot.supportProfile,
  });
  if (!consistencyMatches(snapshot.consistency, sourceConsistency)) {
    consistencyFailure(
      'PowerPoint round-trip source does not match the snapshot',
    );
  }

  if (snapshot.operations.length !== 0) {
    const patched = await patchPptxTextOperations(
      normalized.bytes,
      parsed,
      snapshot.operations,
      limits,
    );
    let outputDocument: Awaited<ReturnType<typeof parse>>;
    try {
      outputDocument = await parse(patched.data, {
        audioMode: 'none',
        errorMode: 'strict',
        imageMode: 'none',
        limits,
        videoMode: 'none',
      });
    } catch (cause) {
      throw new PptxWriteError(
        'verification-failed',
        'PowerPoint text edit output failed strict verification',
        { cause },
      );
    }
    const expectedDocument = applyPptxRoundTripOperationsToPreview(snapshot);
    const outputPreview = createPowerPointRoundTripPreview(outputDocument);
    if (canonicalJson(outputPreview) !== canonicalJson(expectedDocument)) {
      throw new PptxWriteError(
        'verification-failed',
        'PowerPoint text edit output does not match the requested semantics',
      );
    }
    return {
      data: patched.data,
      report: r2Report(
        snapshot,
        patched.copiedPartCount,
        patched.patchedPartCount,
      ),
    };
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

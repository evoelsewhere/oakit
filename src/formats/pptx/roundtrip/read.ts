import { parse } from '../parser';
import type { PptxSceneDocument } from '../scene-types';
import { validatePptxScene } from '../scene-validation';
import { resolvePptxResourceLimits } from '../internal/resource-limits';
import type { PptxInput } from '../types';
import { PptxWriteError } from '../write-error';
import {
  createPptxRoundTripSupportProfile,
  createPptxSnapshotConsistency,
} from './consistency';
import { createPowerPointRoundTripPreview } from './preview';
import {
  detectPptxRoundTripConformance,
  normalizePptxRoundTripInput,
} from './source';
import type {
  PptxRoundTripReadOptions,
  PptxRoundTripRuntimeSource,
  PptxRoundTripSnapshot,
} from './types';

export function assertValidPptxRoundTripPreview(
  document: PptxSceneDocument,
): void {
  const validation = validatePptxScene(document);
  if (!validation.valid) {
    throw new PptxWriteError(
      'invalid-snapshot',
      'PowerPoint semantic preview is not valid for round-trip',
      { issues: validation.issues },
    );
  }
}

export async function readPptxRoundTrip(
  input: PptxInput,
  options: PptxRoundTripReadOptions = {},
): Promise<PptxRoundTripSnapshot> {
  const limits = resolvePptxResourceLimits(options.limits);
  const normalized = await normalizePptxRoundTripInput(input, limits);
  const parsed = await parse(normalized.bytes, {
    audioMode: 'none',
    errorMode: 'strict',
    imageMode: 'none',
    limits,
    videoMode: 'none',
  });
  const conformance = await detectPptxRoundTripConformance(
    normalized.bytes,
    limits,
  );
  const document = createPowerPointRoundTripPreview(parsed);
  assertValidPptxRoundTripPreview(document);

  const source: PptxRoundTripRuntimeSource = {
    byteLength: normalized.byteLength,
    conformance,
    data: normalized.data,
    kind: 'bytes',
    sha256: normalized.sha256,
  };
  const supportProfile = createPptxRoundTripSupportProfile();
  const operations: [] = [];
  const consistency = await createPptxSnapshotConsistency({
    document,
    operations,
    source,
    supportProfile,
  });

  return {
    consistency,
    document,
    format: 'pptx',
    operations,
    schemaVersion: 1,
    source,
    supportProfile,
  };
}

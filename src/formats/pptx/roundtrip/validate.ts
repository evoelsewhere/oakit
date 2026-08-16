import {
  assertPptxInputWithinLimits,
  type ResolvedPptxResourceLimits,
} from '../internal/resource-limits';
import { validatePptxScene } from '../scene-validation';
import { PptxWriteError } from '../write-error';
import {
  PPTX_ROUND_TRIP_CANONICALIZATION_VERSION,
  PPTX_ROUND_TRIP_CAPABILITY_PROFILE_VERSION,
  PPTX_ROUND_TRIP_CONTRACT_VERSION,
  PPTX_ROUND_TRIP_KEY_ALGORITHM_VERSION,
} from './consistency';
import { assertPptxRoundTripDataTree } from './data-tree';
import type { PptxRoundTripSnapshot } from './types';

const ROOT_KEYS = [
  'consistency',
  'document',
  'format',
  'operations',
  'schemaVersion',
  'source',
  'supportProfile',
] as const;
const SOURCE_KEYS = [
  'byteLength',
  'conformance',
  'data',
  'kind',
  'sha256',
] as const;
const SUPPORT_PROFILE_KEYS = [
  'effectiveLevel',
  'id',
  'producerMatrix',
  'version',
] as const;
const CONSISTENCY_KEYS = [
  'canonicalizationVersion',
  'capabilityProfileVersion',
  'contractVersion',
  'hashAlgorithm',
  'keyAlgorithmVersion',
  'operationsSha256',
  'semanticPreviewSha256',
  'sourceManifestSha256',
] as const;
const SHA256_PATTERN = /^[0-9a-f]{64}$/;

function invalidSnapshot(message: string): never {
  throw new PptxWriteError('invalid-snapshot', message);
}

function exactRecord(
  value: unknown,
  keys: readonly string[],
  message: string,
): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    invalidSnapshot(message);
  }
  const names = Object.getOwnPropertyNames(value);
  if (
    names.length !== keys.length ||
    !keys.every((key) => Object.hasOwn(value, key))
  ) {
    invalidSnapshot(message);
  }
  return value as Record<string, unknown>;
}

function assertSha256(
  value: unknown,
  message: string,
): asserts value is string {
  if (typeof value !== 'string' || !SHA256_PATTERN.test(value)) {
    invalidSnapshot(message);
  }
}

function assertLiteral(
  value: unknown,
  expected: string | number,
  message: string,
): void {
  if (value !== expected) invalidSnapshot(message);
}

function validateSource(
  value: unknown,
  limits: ResolvedPptxResourceLimits,
): void {
  const source = exactRecord(
    value,
    SOURCE_KEYS,
    'PowerPoint round-trip snapshot source has an invalid shape',
  );
  assertLiteral(
    source.kind,
    'bytes',
    'PowerPoint round-trip snapshot source kind must be bytes',
  );
  if (!(source.data instanceof Uint8Array) && !(source.data instanceof Blob)) {
    invalidSnapshot(
      'PowerPoint round-trip snapshot source data must be Uint8Array or Blob',
    );
  }
  assertPptxInputWithinLimits(source.data, limits);
  if (
    !Number.isSafeInteger(source.byteLength) ||
    Number(source.byteLength) <= 0
  ) {
    invalidSnapshot(
      'PowerPoint round-trip snapshot source byteLength must be a positive safe integer',
    );
  }
  const actualByteLength =
    source.data instanceof Blob ? source.data.size : source.data.byteLength;
  if (source.byteLength !== actualByteLength) {
    invalidSnapshot(
      'PowerPoint round-trip snapshot source byteLength does not match its data',
    );
  }
  if (
    source.conformance !== 'strict' &&
    source.conformance !== 'transitional' &&
    source.conformance !== 'unknown'
  ) {
    invalidSnapshot(
      'PowerPoint round-trip snapshot source conformance is invalid',
    );
  }
  assertSha256(
    source.sha256,
    'PowerPoint round-trip snapshot source SHA-256 is invalid',
  );
}

function validateSupportProfile(value: unknown): void {
  const profile = exactRecord(
    value,
    SUPPORT_PROFILE_KEYS,
    'PowerPoint round-trip snapshot support profile has an invalid shape',
  );
  assertLiteral(
    profile.effectiveLevel,
    'R0',
    'PowerPoint round-trip snapshot support level must be R0',
  );
  assertLiteral(
    profile.id,
    'pptx-roundtrip-r0',
    'PowerPoint round-trip snapshot support profile id is unsupported',
  );
  assertLiteral(
    profile.version,
    '1',
    'PowerPoint round-trip snapshot support profile version is unsupported',
  );
  if (!Array.isArray(profile.producerMatrix)) {
    invalidSnapshot(
      'PowerPoint round-trip snapshot producer matrix must be an array',
    );
  }
  if (profile.producerMatrix.length !== 0) {
    invalidSnapshot(
      'PowerPoint round-trip R0 snapshot cannot claim producer evidence',
    );
  }
}

function validateConsistency(value: unknown): void {
  const consistency = exactRecord(
    value,
    CONSISTENCY_KEYS,
    'PowerPoint round-trip snapshot consistency has an invalid shape',
  );
  const literals: ReadonlyArray<readonly [unknown, string, string]> = [
    [
      consistency.canonicalizationVersion,
      PPTX_ROUND_TRIP_CANONICALIZATION_VERSION,
      'PowerPoint round-trip snapshot canonicalization version is unsupported',
    ],
    [
      consistency.capabilityProfileVersion,
      PPTX_ROUND_TRIP_CAPABILITY_PROFILE_VERSION,
      'PowerPoint round-trip snapshot capability profile version is unsupported',
    ],
    [
      consistency.contractVersion,
      PPTX_ROUND_TRIP_CONTRACT_VERSION,
      'PowerPoint round-trip snapshot contract version is unsupported',
    ],
    [
      consistency.hashAlgorithm,
      'sha256',
      'PowerPoint round-trip snapshot hash algorithm is unsupported',
    ],
    [
      consistency.keyAlgorithmVersion,
      PPTX_ROUND_TRIP_KEY_ALGORITHM_VERSION,
      'PowerPoint round-trip snapshot key algorithm version is unsupported',
    ],
  ];
  for (const [actual, expected, message] of literals) {
    assertLiteral(actual, expected, message);
  }
  assertSha256(
    consistency.operationsSha256,
    'PowerPoint round-trip snapshot operations SHA-256 is invalid',
  );
  assertSha256(
    consistency.semanticPreviewSha256,
    'PowerPoint round-trip snapshot semantic preview SHA-256 is invalid',
  );
  assertSha256(
    consistency.sourceManifestSha256,
    'PowerPoint round-trip snapshot source manifest SHA-256 is invalid',
  );
}

export function validatePptxRoundTripSnapshot(
  value: unknown,
  limits: ResolvedPptxResourceLimits,
): PptxRoundTripSnapshot {
  assertPptxRoundTripDataTree(value, limits);
  const snapshot = exactRecord(
    value,
    ROOT_KEYS,
    'PowerPoint round-trip snapshot has an invalid root shape',
  );
  assertLiteral(
    snapshot.schemaVersion,
    1,
    'PowerPoint round-trip snapshot schema version is unsupported',
  );
  assertLiteral(
    snapshot.format,
    'pptx',
    'PowerPoint round-trip snapshot format must be pptx',
  );
  if (!Array.isArray(snapshot.operations)) {
    invalidSnapshot(
      'PowerPoint round-trip snapshot operations must be an array',
    );
  }
  if (snapshot.operations.length !== 0) {
    throw new PptxWriteError(
      'unsupported-edit-operation',
      'PowerPoint R0 round-trip does not support edit operations',
    );
  }

  validateSource(snapshot.source, limits);
  validateSupportProfile(snapshot.supportProfile);
  validateConsistency(snapshot.consistency);
  const sceneValidation = validatePptxScene(snapshot.document);
  if (!sceneValidation.valid) {
    throw new PptxWriteError(
      'invalid-snapshot',
      'PowerPoint round-trip snapshot semantic preview is invalid',
      { issues: sceneValidation.issues },
    );
  }

  return snapshot as unknown as PptxRoundTripSnapshot;
}

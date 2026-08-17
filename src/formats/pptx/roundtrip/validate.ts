import {
  assertPptxInputWithinLimits,
  type ResolvedPptxResourceLimits,
} from '../internal/resource-limits';
import { isValidXmlText, validatePptxScene } from '../scene-validation';
import { PptxWriteError } from '../write-error';
import { canonicalJson } from './canonical-json';
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
const REPLACE_TEXT_OPERATION_KEYS = [
  'expectedText',
  'id',
  'kind',
  'targetKey',
  'value',
] as const;
const SET_TRANSFORM_OPERATION_KEYS = [
  'expectedTransform',
  'id',
  'kind',
  'targetKey',
  'value',
] as const;
const TRANSFORM_KEYS = [
  'flipHorizontal',
  'flipVertical',
  'height',
  'rotation',
  'width',
  'x',
  'y',
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

function validateSupportProfile(value: unknown, hasOperations: boolean): void {
  const profile = exactRecord(
    value,
    SUPPORT_PROFILE_KEYS,
    'PowerPoint round-trip snapshot support profile has an invalid shape',
  );
  assertLiteral(
    profile.effectiveLevel,
    hasOperations ? 'R2' : 'R0',
    hasOperations
      ? 'PowerPoint text edit snapshot support level must be R2'
      : 'PowerPoint round-trip snapshot support level must be R0',
  );
  assertLiteral(
    profile.id,
    hasOperations ? 'pptx-roundtrip-text-v1' : 'pptx-roundtrip-r0',
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
      'PowerPoint round-trip snapshot cannot claim unverified producer evidence',
    );
  }
}

function editableRuns(
  value: PptxRoundTripSnapshot['document'],
): Map<string, string> {
  const runs = new Map<string, string>();
  for (const slide of value.slides) {
    for (const element of slide.elements) {
      if (element.type !== 'text') continue;
      for (const paragraph of element.text.paragraphs) {
        for (const child of paragraph.children) {
          if (child.type !== 'run') continue;
          if (runs.has(child.key)) {
            invalidSnapshot(
              'PowerPoint round-trip snapshot contains an ambiguous text target',
            );
          }
          runs.set(child.key, child.text);
        }
      }
    }
  }
  return runs;
}

function editableTransforms(
  value: PptxRoundTripSnapshot['document'],
): Map<string, unknown> {
  const transforms = new Map<string, unknown>();
  for (const slide of value.slides) {
    for (const element of slide.elements) {
      if (element.type !== 'text' || element.resolved.transform === undefined) {
        continue;
      }
      if (transforms.has(element.key)) {
        invalidSnapshot(
          'PowerPoint round-trip snapshot contains an ambiguous transform target',
        );
      }
      transforms.set(element.key, element.resolved.transform);
    }
  }
  return transforms;
}

function validateTransform(
  value: unknown,
  message: string,
): Record<string, unknown> {
  const transform = exactRecord(value, TRANSFORM_KEYS, message);
  for (const key of ['x', 'y', 'width', 'height', 'rotation'] as const) {
    if (
      typeof transform[key] !== 'number' ||
      !Number.isFinite(transform[key])
    ) {
      invalidSnapshot(message);
    }
  }
  if (Number(transform.width) <= 0 || Number(transform.height) <= 0) {
    invalidSnapshot(message);
  }
  for (const key of ['flipHorizontal', 'flipVertical'] as const) {
    if (typeof transform[key] !== 'boolean') invalidSnapshot(message);
  }
  return transform;
}

function validateOperations(
  values: unknown[],
  document: PptxRoundTripSnapshot['document'],
  limits: ResolvedPptxResourceLimits,
): void {
  const runs = editableRuns(document);
  const transforms = editableTransforms(document);
  const ids = new Set<string>();
  const targets = new Set<string>();
  for (const [index, value] of values.entries()) {
    if (value === null || typeof value !== 'object' || Array.isArray(value)) {
      invalidSnapshot(
        `PowerPoint round-trip operation ${index + 1} has an invalid shape`,
      );
    }
    const kind = (value as Record<string, unknown>).kind;
    if (kind !== 'replace-text' && kind !== 'set-transform') {
      invalidSnapshot(
        `PowerPoint round-trip operation ${index + 1} kind is unsupported`,
      );
    }
    const operation = exactRecord(
      value,
      kind === 'replace-text'
        ? REPLACE_TEXT_OPERATION_KEYS
        : SET_TRANSFORM_OPERATION_KEYS,
      `PowerPoint round-trip operation ${index + 1} has an invalid shape`,
    );
    for (const key of ['id', 'targetKey'] as const) {
      if (typeof operation[key] !== 'string') {
        invalidSnapshot(
          `PowerPoint round-trip operation ${index + 1} ${key} must be a string`,
        );
      }
    }
    const id = operation.id as string;
    const targetKey = operation.targetKey as string;
    if (id.length === 0 || ids.has(id)) {
      invalidSnapshot('PowerPoint round-trip operation ids must be unique');
    }
    if (targetKey.length === 0 || targets.has(targetKey)) {
      invalidSnapshot(
        'PowerPoint round-trip text edit targets must be non-empty and unique',
      );
    }
    ids.add(id);
    targets.add(targetKey);
    if (kind === 'set-transform') {
      const expectedTransform = validateTransform(
        operation.expectedTransform,
        `PowerPoint round-trip operation ${index + 1} expectedTransform is invalid`,
      );
      const replacement = validateTransform(
        operation.value,
        `PowerPoint round-trip operation ${index + 1} value is not a valid transform`,
      );
      const sourceTransform = transforms.get(targetKey);
      if (sourceTransform === undefined) {
        invalidSnapshot(
          'PowerPoint round-trip transform target does not exist',
        );
      }
      if (canonicalJson(sourceTransform) !== canonicalJson(expectedTransform)) {
        invalidSnapshot(
          'PowerPoint round-trip transform precondition does not match the preview',
        );
      }
      if (canonicalJson(replacement) === canonicalJson(expectedTransform)) {
        invalidSnapshot(
          'PowerPoint round-trip transform must change the value',
        );
      }
      continue;
    }
    for (const key of ['expectedText', 'value'] as const) {
      if (typeof operation[key] !== 'string') {
        invalidSnapshot(
          `PowerPoint round-trip operation ${index + 1} ${key} must be a string`,
        );
      }
    }
    const expectedText = operation.expectedText as string;
    const replacement = operation.value as string;
    const sourceText = runs.get(targetKey);
    if (sourceText === undefined) {
      invalidSnapshot('PowerPoint round-trip text edit target does not exist');
    }
    if (sourceText !== expectedText) {
      invalidSnapshot(
        'PowerPoint round-trip text edit precondition does not match the preview',
      );
    }
    if (replacement === expectedText) {
      invalidSnapshot('PowerPoint round-trip text edit must change the value');
    }
    if (!isValidXmlText(replacement)) {
      invalidSnapshot(
        'PowerPoint round-trip text edit value is not safe XML text',
      );
    }
    if (
      replacement.length > limits.maxXmlBytes ||
      new TextEncoder().encode(replacement).byteLength > limits.maxXmlBytes
    ) {
      invalidSnapshot(
        'PowerPoint round-trip text edit value exceeds the XML part byte limit',
      );
    }
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

  validateSource(snapshot.source, limits);
  const sceneValidation = validatePptxScene(snapshot.document);
  if (!sceneValidation.valid) {
    throw new PptxWriteError(
      'invalid-snapshot',
      'PowerPoint round-trip snapshot semantic preview is invalid',
      { issues: sceneValidation.issues },
    );
  }
  validateOperations(
    snapshot.operations,
    snapshot.document as PptxRoundTripSnapshot['document'],
    limits,
  );
  validateSupportProfile(
    snapshot.supportProfile,
    snapshot.operations.length !== 0,
  );
  validateConsistency(snapshot.consistency);

  return snapshot as unknown as PptxRoundTripSnapshot;
}

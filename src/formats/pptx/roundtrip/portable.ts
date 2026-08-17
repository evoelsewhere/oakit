import {
  decodeBase64,
  decodedBase64ByteLength,
  encodeBase64,
} from '../../../common/binary/base64';
import {
  resolvePptxResourceLimits,
  type ResolvedPptxResourceLimits,
} from '../internal/resource-limits';
import { PptxWriteError } from '../write-error';
import { assertPptxRoundTripDataTree } from './data-tree';
import { normalizePptxRoundTripInput } from './source';
import type {
  PptxRoundTripPortableJson,
  PptxRoundTripPortableLimits,
  PptxRoundTripSnapshot,
} from './types';
import { validatePptxRoundTripSnapshot } from './validate';
import { writePptxRoundTripWithLimits } from './write';

type PortableLimitName = keyof PptxRoundTripPortableLimits;
type ResolvedPortableLimits = Required<PptxRoundTripPortableLimits>;

export class PptxRoundTripPortableLimitError extends Error {
  readonly actual: number;
  readonly limit: number;
  readonly limitName: PortableLimitName;

  constructor(limitName: PortableLimitName, actual: number, limit: number) {
    super(
      `PowerPoint portable snapshot limit ${limitName} exceeded: ${actual} > ${limit}`,
    );
    this.name = 'PptxRoundTripPortableLimitError';
    this.actual = actual;
    this.limit = limit;
    this.limitName = limitName;
  }
}

export function defaultPptxRoundTripPortableLimits(): ResolvedPortableLimits {
  const maxDecodedBytes = 100 * 1024 * 1024;
  return {
    maxBase64Characters: Math.ceil(maxDecodedBytes / 3) * 4,
    maxDecodedBytes,
  };
}

function assertPositiveSafeInteger(name: PortableLimitName, value: number) {
  if (!Number.isSafeInteger(value) || value <= 0) {
    throw new RangeError(
      `PowerPoint portable snapshot limit ${name} must be a positive integer`,
    );
  }
}

function resolvePortableLimits(
  limits: PptxRoundTripPortableLimits,
): ResolvedPortableLimits {
  const defaults = defaultPptxRoundTripPortableLimits();
  const resolved = {
    maxBase64Characters:
      limits.maxBase64Characters ?? defaults.maxBase64Characters,
    maxDecodedBytes: limits.maxDecodedBytes ?? defaults.maxDecodedBytes,
  };
  assertPositiveSafeInteger(
    'maxBase64Characters',
    resolved.maxBase64Characters,
  );
  assertPositiveSafeInteger('maxDecodedBytes', resolved.maxDecodedBytes);
  return resolved;
}

function resolveInputLimits(maxInputBytes: number): ResolvedPptxResourceLimits {
  const limits = resolvePptxResourceLimits();
  limits.maxInputBytes = maxInputBytes;
  return limits;
}

function assertPortableLimit(
  name: PortableLimitName,
  actual: number,
  limit: number,
): void {
  if (actual > limit) {
    throw new PptxRoundTripPortableLimitError(name, actual, limit);
  }
}

function invalidPortableSnapshot(message: string, cause?: unknown): never {
  throw new PptxWriteError('invalid-snapshot', message, { cause });
}

function rootKeys(): readonly string[] {
  return [
    'consistency',
    'document',
    'format',
    'operations',
    'schemaVersion',
    'source',
    'supportProfile',
  ];
}

function sourceKeys(): readonly string[] {
  return ['byteLength', 'conformance', 'kind', 'packageBase64', 'sha256'];
}

function exactRecord(
  value: unknown,
  keys: readonly string[],
  message: string,
): Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) {
    invalidPortableSnapshot(message);
  }
  const names = Object.getOwnPropertyNames(value);
  if (
    names.length !== keys.length ||
    !keys.every((key) => Object.hasOwn(value, key))
  ) {
    invalidPortableSnapshot(message);
  }
  return value as Record<string, unknown>;
}

function encodedBase64Length(byteLength: number): number {
  return Math.ceil(byteLength / 3) * 4;
}

function decodedLength(packageBase64: string): number {
  try {
    return decodedBase64ByteLength(packageBase64);
  } catch (cause) {
    invalidPortableSnapshot(
      'PowerPoint portable snapshot packageBase64 must be canonical Base64',
      cause,
    );
  }
}

export async function serializePptxRoundTripJson(
  value: PptxRoundTripSnapshot,
  portableLimits: PptxRoundTripPortableLimits = {},
): Promise<PptxRoundTripPortableJson> {
  const limits = resolvePortableLimits(portableLimits);
  const validationLimits = resolveInputLimits(Number.MAX_SAFE_INTEGER);
  const validated = validatePptxRoundTripSnapshot(value, validationLimits);
  const snapshot = structuredClone(validated);
  assertPortableLimit(
    'maxDecodedBytes',
    snapshot.source.byteLength,
    limits.maxDecodedBytes,
  );
  assertPortableLimit(
    'maxBase64Characters',
    encodedBase64Length(snapshot.source.byteLength),
    limits.maxBase64Characters,
  );

  const runtimeLimits = resolveInputLimits(limits.maxDecodedBytes);
  const normalized = await normalizePptxRoundTripInput(
    snapshot.source.data,
    runtimeLimits,
  );
  snapshot.source.data = normalized.bytes;
  await writePptxRoundTripWithLimits(snapshot, runtimeLimits);
  const packageBase64 = encodeBase64(normalized.bytes);
  return {
    consistency: snapshot.consistency,
    document: snapshot.document,
    format: 'pptx',
    operations: snapshot.operations,
    schemaVersion: 1,
    source: {
      byteLength: snapshot.source.byteLength,
      conformance: snapshot.source.conformance,
      kind: 'base64',
      packageBase64,
      sha256: snapshot.source.sha256,
    },
    supportProfile: snapshot.supportProfile,
  };
}

export async function parsePptxRoundTripJson(
  value: unknown,
  portableLimits: PptxRoundTripPortableLimits = {},
): Promise<PptxRoundTripSnapshot> {
  const limits = resolvePortableLimits(portableLimits);
  const runtimeLimits = resolveInputLimits(limits.maxDecodedBytes);
  assertPptxRoundTripDataTree(value, runtimeLimits);
  const root = exactRecord(
    value,
    rootKeys(),
    'PowerPoint portable snapshot has an invalid root shape',
  );
  const source = exactRecord(
    root.source,
    sourceKeys(),
    'PowerPoint portable snapshot source has an invalid shape',
  );
  if (source.kind !== 'base64') {
    invalidPortableSnapshot(
      'PowerPoint portable snapshot source kind must be base64',
    );
  }
  if (typeof source.packageBase64 !== 'string') {
    invalidPortableSnapshot(
      'PowerPoint portable snapshot packageBase64 must be a string',
    );
  }
  assertPortableLimit(
    'maxBase64Characters',
    source.packageBase64.length,
    limits.maxBase64Characters,
  );
  const byteLength = decodedLength(source.packageBase64);
  assertPortableLimit('maxDecodedBytes', byteLength, limits.maxDecodedBytes);
  if (
    !Number.isSafeInteger(source.byteLength) ||
    Number(source.byteLength) <= 0
  ) {
    invalidPortableSnapshot(
      'PowerPoint portable snapshot source byteLength must be a positive safe integer',
    );
  }
  if (source.byteLength !== byteLength) {
    invalidPortableSnapshot(
      'PowerPoint portable snapshot source byteLength does not match packageBase64',
    );
  }

  const owned = structuredClone(root) as unknown as PptxRoundTripPortableJson;
  const runtimeValue = {
    consistency: owned.consistency,
    document: owned.document,
    format: owned.format,
    operations: owned.operations,
    schemaVersion: owned.schemaVersion,
    source: {
      byteLength: owned.source.byteLength,
      conformance: owned.source.conformance,
      data: decodeBase64(owned.source.packageBase64),
      kind: 'bytes',
      sha256: owned.source.sha256,
    },
    supportProfile: owned.supportProfile,
  };
  const snapshot = validatePptxRoundTripSnapshot(runtimeValue, runtimeLimits);
  await writePptxRoundTripWithLimits(snapshot, runtimeLimits);
  return snapshot;
}

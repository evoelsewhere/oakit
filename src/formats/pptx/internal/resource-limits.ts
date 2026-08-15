import type JSZip from 'jszip';

import type { PptxDiagnostic, PptxInput, PptxResourceLimits } from '../types';

export type ResolvedPptxResourceLimits = Required<PptxResourceLimits>;

export function defaultPptxResourceLimits(): ResolvedPptxResourceLimits {
  const mebibyte = 1024 * 1024;
  return {
    maxEntries: 10_000,
    maxInputBytes: 100 * mebibyte,
    maxMediaBytes: 64 * mebibyte,
    maxPartBytes: 64 * mebibyte,
    maxSlides: 1_000,
    maxTotalUncompressedBytes: 256 * mebibyte,
    maxTotalXmlNodes: 1_000_000,
    maxXmlBytes: 16 * mebibyte,
    maxXmlDepth: 128,
    maxXmlNodes: 250_000,
  };
}

interface CompressedEntryData {
  uncompressedSize?: unknown;
}

interface SizedZipObject extends JSZip.JSZipObject {
  _data?: CompressedEntryData | Promise<unknown>;
}

export class PptxResourceLimitError extends Error {
  readonly actual: number;
  readonly limit: number;
  readonly limitName: keyof PptxResourceLimits;
  readonly part?: string;

  constructor(
    limitName: keyof PptxResourceLimits,
    actual: number,
    limit: number,
    part?: string,
  ) {
    const location = part ? ` for ${part}` : '';
    super(
      `PPTX resource limit ${limitName} exceeded${location}: ${actual} > ${limit}`,
    );
    this.name = 'PptxResourceLimitError';
    this.actual = actual;
    this.limit = limit;
    this.limitName = limitName;
    if (part) this.part = part;
  }
}

function assertPositiveSafeInteger(
  name: keyof PptxResourceLimits,
  value: number,
): void {
  if (!Number.isSafeInteger(value) || value <= 0) {
    throw new RangeError(
      `PPTX resource limit ${name} must be a positive integer`,
    );
  }
}

export function resolvePptxResourceLimits(
  limits: PptxResourceLimits = {},
): ResolvedPptxResourceLimits {
  const resolved: ResolvedPptxResourceLimits = {
    ...defaultPptxResourceLimits(),
    ...limits,
  };
  for (const [name, value] of Object.entries(resolved)) {
    assertPositiveSafeInteger(name as keyof PptxResourceLimits, value);
  }
  if (resolved.maxXmlBytes > resolved.maxPartBytes) {
    throw new RangeError('maxXmlBytes cannot exceed maxPartBytes');
  }
  if (resolved.maxMediaBytes > resolved.maxPartBytes) {
    throw new RangeError('maxMediaBytes cannot exceed maxPartBytes');
  }
  return resolved;
}

function inputByteLength(input: PptxInput): number {
  if (input instanceof ArrayBuffer) return input.byteLength;
  if (ArrayBuffer.isView(input)) return input.byteLength;
  return input.size;
}

export function assertPptxInputWithinLimits(
  input: PptxInput,
  limits: ResolvedPptxResourceLimits,
): void {
  const size = inputByteLength(input);
  if (size > limits.maxInputBytes) {
    throw new PptxResourceLimitError(
      'maxInputBytes',
      size,
      limits.maxInputBytes,
    );
  }
}

function declaredUncompressedSize(file: JSZip.JSZipObject): number {
  const data = (file as SizedZipObject)._data;
  if (data instanceof Promise) return 0;
  const size = data?.uncompressedSize;
  const numericSize = Number(size);
  if (!Number.isSafeInteger(size) || numericSize !== Math.abs(numericSize)) {
    throw new Error(
      `Unable to validate expanded size for ZIP part ${file.name}`,
    );
  }
  return numericSize;
}

export function assertPptxArchiveWithinLimits(
  zip: JSZip,
  limits: ResolvedPptxResourceLimits,
): void {
  const entries = Object.values(zip.files).filter((file) => !file.dir);
  if (entries.length > limits.maxEntries) {
    throw new PptxResourceLimitError(
      'maxEntries',
      entries.length,
      limits.maxEntries,
    );
  }

  let totalUncompressedBytes = 0;
  for (const entry of entries) {
    const size = declaredUncompressedSize(entry);
    if (size > limits.maxPartBytes) {
      throw new PptxResourceLimitError(
        'maxPartBytes',
        size,
        limits.maxPartBytes,
        entry.name,
      );
    }
    totalUncompressedBytes += size;
    if (
      !Number.isSafeInteger(totalUncompressedBytes) ||
      totalUncompressedBytes > limits.maxTotalUncompressedBytes
    ) {
      throw new PptxResourceLimitError(
        'maxTotalUncompressedBytes',
        totalUncompressedBytes,
        limits.maxTotalUncompressedBytes,
      );
    }
  }
}

export function resourceLimitDiagnostic(
  error: PptxResourceLimitError,
): PptxDiagnostic {
  return {
    code: 'resource-limit-exceeded',
    message: error.message,
    severity: 'error',
    ...(error.part ? { part: error.part } : {}),
  };
}

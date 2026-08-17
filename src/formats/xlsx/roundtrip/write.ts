import { createXlsxCapabilityManifest } from './capability';
import { resolveXlsxWriteLimits } from './write-limits';
import type {
  XlsxRoundTripSnapshot,
  XlsxWriteOptions,
  XlsxWriteResult,
} from './types';
import { verifyXlsxRoundTripSnapshot } from './verify';

function assertWriteOptions(options: XlsxWriteOptions): void {
  if (
    Object.prototype.toString.call(options) !== '[object Object]' ||
    Object.getPrototypeOf(options) !== Object.prototype
  ) {
    throw new TypeError('XLSX write options must be a plain object');
  }
  for (const key of Object.keys(options)) {
    if (
      key !== 'acknowledgeOpaqueContent' &&
      key !== 'limits' &&
      key !== 'minimumEditedFidelity' &&
      key !== 'readerLimits'
    ) {
      throw new TypeError(`Unknown XLSX write option ${key}`);
    }
  }
  if (
    options.acknowledgeOpaqueContent !== undefined &&
    typeof options.acknowledgeOpaqueContent !== 'boolean'
  ) {
    throw new TypeError('XLSX acknowledgeOpaqueContent must be boolean');
  }
  if (
    options.minimumEditedFidelity !== undefined &&
    options.minimumEditedFidelity !== 'R1' &&
    options.minimumEditedFidelity !== 'R2' &&
    options.minimumEditedFidelity !== 'R3'
  ) {
    throw new TypeError('XLSX minimumEditedFidelity is invalid');
  }
}

export async function writeXlsxRoundTrip(
  value: XlsxRoundTripSnapshot,
  options: XlsxWriteOptions = {},
): Promise<XlsxWriteResult> {
  assertWriteOptions(options);
  const writeLimits = resolveXlsxWriteLimits(options.limits);
  const verified = await verifyXlsxRoundTripSnapshot(
    value,
    options,
    writeLimits,
  );
  return {
    data: verified.bytes,
    report: {
      diagnostics: [],
      level: 'R0',
      outputSha256: verified.snapshot.source.sha256,
      parts: verified.graph.parts.map((part) => ({
        byteLength: part.byteLength,
        disposition: 'copy',
        name: part.name,
        sha256: part.sha256,
      })),
      sourceSha256: verified.snapshot.source.sha256,
      supportProfile: createXlsxCapabilityManifest(),
    },
  };
}

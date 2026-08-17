import { resolveXlsxWriteLimits } from './write-limits';
import type { XlsxRoundTripSnapshot, XlsxWriteOptions } from './types';
import { verifyXlsxRoundTripSnapshot } from './verify';

export async function validateXlsxRoundTripJson(
  value: unknown,
  options: XlsxWriteOptions = {},
): Promise<XlsxRoundTripSnapshot> {
  const writeLimits = resolveXlsxWriteLimits(options.limits);
  const verified = await verifyXlsxRoundTripSnapshot(
    value,
    options,
    writeLimits,
  );
  return structuredClone(verified.snapshot);
}

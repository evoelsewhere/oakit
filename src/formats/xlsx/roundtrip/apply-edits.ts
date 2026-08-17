import { XlsxWriteError } from './errors';
import type {
  XlsxEditOperation,
  XlsxRoundTripSnapshot,
  XlsxWriteOptions,
} from './types';
import { validateXlsxRoundTripJson } from './validate-json';

export async function applyXlsxEdits(
  snapshot: XlsxRoundTripSnapshot,
  operations: readonly XlsxEditOperation[],
  options: XlsxWriteOptions = {},
): Promise<XlsxRoundTripSnapshot> {
  if (!Array.isArray(operations)) {
    throw new TypeError('XLSX edit operations must be an array');
  }
  if (operations.length !== 0) {
    const first = (operations as readonly XlsxEditOperation[])[0]!;
    throw new XlsxWriteError(
      'unsupported-edit-operation',
      'The XLSX R0 capability profile does not support edit operations',
      { operationId: first.operationId },
    );
  }
  return validateXlsxRoundTripJson(snapshot, options);
}

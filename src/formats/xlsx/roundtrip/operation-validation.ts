import type { ResolvedXlsxResourceLimits } from '../internal/resource-limits';
import { parseXlsxCellReference } from '../internal/cell-reference';
import type { XlsxCellValue, XlsxStyle } from '../types';
import { canonicalXlsxJson } from './canonical-json';
import { XlsxWriteError } from './errors';
import type { ResolvedXlsxWriteLimits, XlsxEditOperation } from './types';
import { writeLimitFailure } from './write-limits';

export type XlsxCellEditOperation = Extract<
  XlsxEditOperation,
  { kind: 'clear-cell' | 'set-cell' | 'set-cell-style' }
>;

const ERROR_CODES = new Set([
  '#BLOCKED!',
  '#BUSY!',
  '#CALC!',
  '#CONNECT!',
  '#DIV/0!',
  '#FIELD!',
  '#GETTING_DATA',
  '#N/A',
  '#NAME?',
  '#NULL!',
  '#NUM!',
  '#REF!',
  '#SPILL!',
  '#UNKNOWN!',
  '#VALUE!',
]);
const KNOWN_OPERATIONS = new Set([
  'add-worksheet',
  'delete-columns',
  'delete-rows',
  'delete-worksheet',
  'insert-columns',
  'insert-rows',
  'rename-worksheet',
  'set-column',
  'set-hyperlink',
  'set-row',
]);
const OPERATION_ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$/u;
const SHA256_PATTERN = /^[0-9a-f]{64}$/u;
const SHEET_KEY_PATTERN = /^xlsx:sheet:[0-9a-f]{32}$/u;
const STYLE_KEYS = [
  'alignment',
  'border',
  'checkbox',
  'fill',
  'font',
  'numberFormat',
  'protection',
] as const;

function invalid(message: string, operationId?: string): never {
  throw new XlsxWriteError('invalid-roundtrip-json', message, {
    ...(operationId === undefined ? {} : { operationId }),
  });
}

function unsupported(
  message: string,
  operationId: string | undefined,
  featureClass?: string,
): never {
  throw new XlsxWriteError('unsupported-edit-operation', message, {
    ...(featureClass === undefined ? {} : { featureClass }),
    ...(operationId === undefined ? {} : { operationId }),
  });
}

function exactKeys(
  value: Record<string, unknown>,
  required: readonly string[],
  optional?: readonly string[],
): boolean {
  const names = Object.keys(value);
  if (optional === undefined) {
    return (
      names.length === required.length &&
      required.every((key) => Object.hasOwn(value, key))
    );
  }
  return (
    required.every((key) => Object.hasOwn(value, key)) &&
    names.every((key) => required.includes(key) || optional.includes(key))
  );
}

function plainRecord(value: unknown): value is Record<string, unknown> {
  if (value === null || value === undefined) return false;
  return Object.getPrototypeOf(value) === Object.prototype;
}

function operationId(value: unknown): string {
  if (typeof value !== 'string' || !OPERATION_ID_PATTERN.test(value)) {
    invalid('XLSX operation ID is invalid');
  }
  return value;
}

function validateCommon(
  operation: Record<string, unknown>,
  id: string,
): { cell: string; operationId: string; sheetKey: string } {
  if (
    typeof operation.sheetKey !== 'string' ||
    !SHEET_KEY_PATTERN.test(operation.sheetKey)
  ) {
    invalid('XLSX operation sheet key is invalid', id);
  }
  const parsed = parseXlsxCellReference(operation.cell);
  if (
    !parsed ||
    parsed.absoluteColumn ||
    parsed.absoluteRow ||
    operation.cell !== parsed.address
  ) {
    invalid('XLSX operation cell reference is invalid', id);
  }
  if (
    operation.ifMatch !== undefined &&
    (typeof operation.ifMatch !== 'string' ||
      !SHA256_PATTERN.test(operation.ifMatch))
  ) {
    invalid('XLSX operation precondition hash is invalid', id);
  }
  return {
    cell: parsed.address,
    operationId: id,
    sheetKey: operation.sheetKey,
  };
}

function validateTextValue(
  value: Record<string, unknown>,
  id: string,
): XlsxCellValue {
  if (!exactKeys(value, ['kind', 'text'], ['runs'])) {
    invalid('XLSX text cell value shape is invalid', id);
  }
  if (typeof value.text !== 'string') {
    invalid('XLSX text cell value is invalid', id);
  }
  if (value.runs !== undefined) {
    unsupported(
      'XLSX cell editing does not yet support rich text runs',
      id,
      'rich-text',
    );
  }
  return { kind: 'text', text: value.text };
}

function validateCellValue(value: unknown, id: string): XlsxCellValue {
  if (!plainRecord(value)) {
    invalid('XLSX cell value shape is invalid', id);
  }
  const record = value;
  if (record.kind === 'text') return validateTextValue(record, id);
  if (record.kind === 'number') {
    if (
      !exactKeys(record, ['kind', 'value']) ||
      typeof record.value !== 'number' ||
      !Number.isFinite(record.value)
    ) {
      invalid('XLSX number cell value is invalid', id);
    }
    return { kind: 'number', value: record.value === 0 ? 0 : record.value };
  }
  if (record.kind === 'boolean') {
    if (
      !exactKeys(record, ['kind', 'value']) ||
      typeof record.value !== 'boolean'
    ) {
      invalid('XLSX boolean cell value is invalid', id);
    }
    return { kind: 'boolean', value: record.value };
  }
  if (record.kind === 'error') {
    if (
      !exactKeys(record, ['code', 'kind']) ||
      typeof record.code !== 'string' ||
      !ERROR_CODES.has(record.code)
    ) {
      invalid('XLSX error cell value is invalid', id);
    }
    return { code: record.code, kind: 'error' };
  }
  if (record.kind === 'date') {
    unsupported(
      'XLSX cell editing does not yet support date values',
      id,
      'date-value',
    );
  }
  invalid('XLSX cell value kind is invalid', id);
}

function validateContent(
  value: unknown,
  id: string,
  readerLimits: ResolvedXlsxResourceLimits,
): Extract<XlsxEditOperation, { kind: 'set-cell' }>['content'] {
  if (!plainRecord(value)) {
    invalid('XLSX set-cell content shape is invalid', id);
  }
  const content = value;
  if (content.kind === 'formula') {
    if (
      !exactKeys(content, ['expression', 'kind']) ||
      typeof content.expression !== 'string' ||
      content.expression.length === 0 ||
      content.expression.startsWith('=')
    ) {
      invalid('XLSX set-cell formula is invalid', id);
    }
    if (content.expression.length > readerLimits.maxFormulaCharacters) {
      throw new XlsxWriteError(
        'resource-limit-exceeded',
        'XLSX operation formula exceeds its character limit',
        {
          actual: content.expression.length,
          limit: readerLimits.maxFormulaCharacters,
          limitName: 'maxFormulaCharacters',
          operationId: id,
        },
      );
    }
    return { expression: content.expression, kind: 'formula' };
  }
  if (content.kind === 'value') {
    if (!exactKeys(content, ['kind', 'value'])) {
      invalid('XLSX set-cell value content shape is invalid', id);
    }
    return { kind: 'value', value: validateCellValue(content.value, id) };
  }
  invalid('XLSX set-cell content kind is invalid', id);
}

function validateExistingStyle(value: unknown, id: string): XlsxStyle {
  if (!plainRecord(value) || !exactKeys(value, [], STYLE_KEYS)) {
    invalid('XLSX set-cell-style style shape is invalid', id);
  }
  return structuredClone(value);
}

function operationBytes(operation: unknown): number {
  return new TextEncoder().encode(canonicalXlsxJson(operation)).byteLength;
}

export function validateXlsxCellOperations(
  value: unknown,
  writeLimits: ResolvedXlsxWriteLimits,
  readerLimits: ResolvedXlsxResourceLimits,
): XlsxCellEditOperation[] {
  if (!Array.isArray(value)) {
    invalid('XLSX round-trip operations must be an array');
  }
  if (value.length > writeLimits.maxOperations) {
    writeLimitFailure('maxOperations', value.length, writeLimits.maxOperations);
  }
  const ids = new Set<string>();
  const operations: XlsxCellEditOperation[] = [];
  for (const candidate of value) {
    if (!plainRecord(candidate)) {
      invalid('XLSX operation shape is invalid');
    }
    const operation = candidate;
    const id = operationId(operation.operationId);
    if (ids.has(id)) {
      invalid('XLSX operation IDs must be unique', id);
    }
    ids.add(id);
    if (operation.kind === 'clear-cell') {
      if (
        !exactKeys(
          operation,
          ['cell', 'kind', 'operationId', 'sheetKey'],
          ['ifMatch'],
        )
      ) {
        invalid('XLSX clear-cell operation shape is invalid', id);
      }
      const common = validateCommon(operation, id);
      operations.push({
        ...common,
        ...(operation.ifMatch === undefined
          ? {}
          : { ifMatch: operation.ifMatch as string }),
        kind: 'clear-cell',
      });
      continue;
    }
    if (operation.kind === 'set-cell') {
      if (
        !exactKeys(
          operation,
          ['cell', 'content', 'kind', 'operationId', 'sheetKey'],
          ['ifMatch'],
        )
      ) {
        invalid('XLSX set-cell operation shape is invalid', id);
      }
      const common = validateCommon(operation, id);
      const content = validateContent(operation.content, id, readerLimits);
      operations.push({
        ...common,
        content,
        ...(operation.ifMatch === undefined
          ? {}
          : { ifMatch: operation.ifMatch as string }),
        kind: 'set-cell',
      });
      continue;
    }
    if (operation.kind === 'set-cell-style') {
      if (
        !exactKeys(
          operation,
          ['cell', 'kind', 'operationId', 'sheetKey', 'style'],
          ['ifMatch'],
        )
      ) {
        invalid('XLSX set-cell-style operation shape is invalid', id);
      }
      const common = validateCommon(operation, id);
      operations.push({
        ...common,
        ...(operation.ifMatch === undefined
          ? {}
          : { ifMatch: operation.ifMatch as string }),
        kind: 'set-cell-style',
        style: validateExistingStyle(operation.style, id),
      });
      continue;
    }
    if (
      typeof operation.kind === 'string' &&
      KNOWN_OPERATIONS.has(operation.kind)
    ) {
      unsupported(
        `XLSX operation ${operation.kind} is not supported by this profile`,
        id,
        operation.kind,
      );
    }
    invalid('XLSX operation kind is invalid', id);
  }
  let totalBytes = 0;
  let totalFormulaCharacters = 0;
  let totalTextCharacters = 0;
  for (const operation of operations) {
    const bytes = operationBytes(operation);
    if (bytes > writeLimits.maxOperationBytes) {
      throw new XlsxWriteError(
        'resource-limit-exceeded',
        'XLSX operation exceeds its byte limit',
        {
          actual: bytes,
          limit: writeLimits.maxOperationBytes,
          limitName: 'maxOperationBytes',
          operationId: operation.operationId,
        },
      );
    }
    totalBytes += bytes;
    if (totalBytes > writeLimits.maxTotalOperationBytes) {
      throw new XlsxWriteError(
        'resource-limit-exceeded',
        'XLSX operations exceed their total byte limit',
        {
          actual: totalBytes,
          limit: writeLimits.maxTotalOperationBytes,
          limitName: 'maxTotalOperationBytes',
          operationId: operation.operationId,
        },
      );
    }
    if (operation.kind === 'set-cell' && operation.content.kind === 'formula') {
      totalFormulaCharacters += operation.content.expression.length;
      if (totalFormulaCharacters > readerLimits.maxTotalFormulaCharacters) {
        throw new XlsxWriteError(
          'resource-limit-exceeded',
          'XLSX operations exceed their total formula character limit',
          {
            actual: totalFormulaCharacters,
            limit: readerLimits.maxTotalFormulaCharacters,
            limitName: 'maxTotalFormulaCharacters',
            operationId: operation.operationId,
          },
        );
      }
    }
    if (
      operation.kind === 'set-cell' &&
      operation.content.kind === 'value' &&
      operation.content.value.kind === 'text'
    ) {
      totalTextCharacters += operation.content.value.text.length;
      if (totalTextCharacters > readerLimits.maxTextCharacters) {
        throw new XlsxWriteError(
          'resource-limit-exceeded',
          'XLSX operations exceed their text character limit',
          {
            actual: totalTextCharacters,
            limit: readerLimits.maxTextCharacters,
            limitName: 'maxTextCharacters',
            operationId: operation.operationId,
          },
        );
      }
    }
  }
  return operations;
}

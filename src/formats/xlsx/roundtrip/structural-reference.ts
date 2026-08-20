import { xlsxColumnName } from '../internal/cell-reference';
import type { XlsxRange } from '../types';

export interface XlsxStructuralReferenceOperation {
  count: number;
  index: number;
  kind: 'delete-columns' | 'delete-rows' | 'insert-columns' | 'insert-rows';
}

function transformCoordinate(
  value: number,
  operation: XlsxStructuralReferenceOperation,
): number | null {
  const insertion = operation.kind.startsWith('insert-');
  if (insertion) {
    return value < operation.index ? value : value + operation.count;
  }
  const end = operation.index + operation.count - 1;
  if (value < operation.index) return value;
  if (value <= end) return null;
  return value - operation.count;
}

function transformInterval(
  start: number,
  end: number,
  operation: XlsxStructuralReferenceOperation,
): readonly [number, number] | null {
  if (operation.kind.startsWith('insert-')) {
    return [
      transformCoordinate(start, operation)!,
      transformCoordinate(end, operation)!,
    ];
  }
  const deletedEnd = operation.index + operation.count - 1;
  if (end < operation.index) return [start, end];
  if (start > deletedEnd) {
    return [start - operation.count, end - operation.count];
  }
  const outputStart = Math.min(start, operation.index);
  const deletedInside = Math.min(operation.count, end - operation.index + 1);
  const outputEnd = end - deletedInside;
  return outputStart > outputEnd ? null : [outputStart, outputEnd];
}

export function transformXlsxStructuralCell(
  row: number,
  column: number,
  operation: XlsxStructuralReferenceOperation,
): { address: string; column: number; row: number } | null {
  const rowOperation =
    operation.kind === 'delete-rows' || operation.kind === 'insert-rows';
  const columnOperation =
    operation.kind === 'delete-columns' || operation.kind === 'insert-columns';
  const transformedRow = rowOperation
    ? transformCoordinate(row, operation)
    : row;
  const transformedColumn = columnOperation
    ? transformCoordinate(column, operation)
    : column;
  if (transformedRow === null || transformedColumn === null) return null;
  return {
    address: `${xlsxColumnName(transformedColumn)!}${transformedRow}`,
    column: transformedColumn,
    row: transformedRow,
  };
}

export function transformXlsxStructuralRange(
  range: XlsxRange,
  operation: XlsxStructuralReferenceOperation,
): XlsxRange | null {
  const rowOperation =
    operation.kind === 'delete-rows' || operation.kind === 'insert-rows';
  const columnOperation =
    operation.kind === 'delete-columns' || operation.kind === 'insert-columns';
  const rows = rowOperation
    ? transformInterval(range.start.row, range.end.row, operation)
    : ([range.start.row, range.end.row] as const);
  const columns = columnOperation
    ? transformInterval(range.start.column, range.end.column, operation)
    : ([range.start.column, range.end.column] as const);
  if (rows === null || columns === null) return null;
  const startAddress = `${xlsxColumnName(columns[0])!}${rows[0]}`;
  const endAddress = `${xlsxColumnName(columns[1])!}${rows[1]}`;
  return {
    end: { column: columns[1], row: rows[1] },
    reference:
      startAddress === endAddress
        ? startAddress
        : `${startAddress}:${endAddress}`,
    start: { column: columns[0], row: rows[0] },
  };
}

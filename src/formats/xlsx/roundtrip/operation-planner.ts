import { parseXlsxCellReference } from '../internal/cell-reference';
import type { ResolvedXlsxResourceLimits } from '../internal/resource-limits';
import type {
  XlsxCell,
  XlsxColumnRange,
  XlsxHyperlink,
  XlsxRow,
  XlsxWorksheet,
} from '../types';
import { canonicalXlsxJson } from './canonical-json';
import { canonicalXlsxSha256 } from './digest';
import { XlsxWriteError } from './errors';
import {
  type XlsxCellEditOperation,
  validateXlsxCellOperations,
} from './operation-validation';
import type {
  ResolvedXlsxWriteLimits,
  XlsxRoundTripDocument,
  XlsxRoundTripSheet,
} from './types';

interface XlsxOperationImpactBase {
  operationId: string;
  sheetKey: string;
}

export type XlsxCellOperationImpact =
  | (XlsxOperationImpactBase & {
      cell: string;
      kind: Extract<
        XlsxCellEditOperation['kind'],
        'clear-cell' | 'set-cell' | 'set-cell-style' | 'set-hyperlink'
      >;
    })
  | (XlsxOperationImpactBase & {
      kind: 'set-column';
      range: string;
    })
  | (XlsxOperationImpactBase & { kind: 'set-row'; range: string });

export interface XlsxCellOperationPlan {
  document: XlsxRoundTripDocument;
  impacts: XlsxCellOperationImpact[];
  operations: XlsxCellEditOperation[];
  stateHash: string;
}

function cloneDocument(document: XlsxRoundTripDocument): XlsxRoundTripDocument {
  return JSON.parse(JSON.stringify(document)) as XlsxRoundTripDocument;
}

function operationFailure(
  code:
    | 'operation-precondition-failed'
    | 'preservation-conflict'
    | 'unsupported-edit-operation',
  message: string,
  operation: XlsxCellEditOperation,
  featureClass?: string,
): never {
  throw new XlsxWriteError(code, message, {
    ...('cell' in operation ? { cell: operation.cell } : {}),
    ...(featureClass === undefined ? {} : { featureClass }),
    operationId: operation.operationId,
    ...('row' in operation
      ? { range: String(operation.row) }
      : 'start' in operation
        ? { range: `${operation.start}:${operation.end}` }
        : {}),
    sheetKey: operation.sheetKey,
  });
}

function isCellOperation(
  operation: XlsxCellEditOperation,
): operation is Extract<XlsxCellEditOperation, { cell: string }> {
  return 'cell' in operation;
}

function resolveWorksheet(
  document: XlsxRoundTripDocument,
  operation: XlsxCellEditOperation,
): XlsxWorksheet & { key: string } {
  const sheet = document.sheets.find(
    (candidate) => candidate.key === operation.sheetKey,
  );
  if (!sheet) {
    operationFailure(
      'preservation-conflict',
      'XLSX operation sheet key does not exist in the snapshot',
      operation,
      'worksheet',
    );
  }
  if (sheet.kind !== 'worksheet') {
    operationFailure(
      'preservation-conflict',
      isCellOperation(operation)
        ? 'XLSX cell operation cannot target a chart sheet'
        : 'XLSX row or column operation cannot target a chart sheet',
      operation,
      'chart-sheet',
    );
  }
  return sheet;
}

function resolveRow(
  sheet: XlsxWorksheet,
  operation: Extract<XlsxCellEditOperation, { kind: 'set-row' }>,
): XlsxRow {
  const row = sheet.rows.find((candidate) => candidate.index === operation.row);
  if (!row) {
    operationFailure(
      'preservation-conflict',
      'XLSX set-row operation requires an existing explicit source row',
      operation,
      'missing-row',
    );
  }
  return row;
}

function resolveColumn(
  sheet: XlsxWorksheet,
  operation: Extract<XlsxCellEditOperation, { kind: 'set-column' }>,
): XlsxColumnRange {
  const column = sheet.columns.find(
    (candidate) =>
      candidate.start === operation.start && candidate.end === operation.end,
  );
  if (!column) {
    operationFailure(
      'preservation-conflict',
      'XLSX set-column operation requires an existing exact source range',
      operation,
      'missing-column-range',
    );
  }
  return column;
}

function resolveCell(
  sheet: XlsxWorksheet,
  operation: Extract<XlsxCellEditOperation, { cell: string }>,
): XlsxCell {
  const reference = parseXlsxCellReference(operation.cell)!;
  const row = sheet.rows.find((candidate) => candidate.index === reference.row);
  const cell = row?.cells.find(
    (candidate) => candidate.column === reference.column,
  );
  if (!cell) {
    operationFailure(
      'preservation-conflict',
      'XLSX cell operation requires an existing explicit source cell',
      operation,
      'missing-cell',
    );
  }
  const merged = sheet.mergedRanges.find(
    (range) =>
      reference.row >= range.start.row &&
      reference.row <= range.end.row &&
      reference.column >= range.start.column &&
      reference.column <= range.end.column,
  );
  if (
    merged &&
    (merged.start.row !== reference.row ||
      merged.start.column !== reference.column)
  ) {
    operationFailure(
      'preservation-conflict',
      'XLSX cell operation cannot target a non-anchor merged cell',
      operation,
      'merged-cell',
    );
  }
  return cell;
}

export function xlsxCellTargetState(
  sheet: XlsxRoundTripSheet,
  cell: XlsxCell,
): { cell: XlsxCell; hyperlink?: XlsxHyperlink; sheetKey: string } {
  const hyperlink =
    sheet.kind === 'worksheet'
      ? sheet.hyperlinks.find(
          (candidate) => candidate.range.reference === cell.address,
        )
      : undefined;
  return {
    cell,
    ...(hyperlink === undefined ? {} : { hyperlink }),
    sheetKey: sheet.key,
  };
}

export function xlsxRowTargetState(
  sheet: XlsxRoundTripSheet,
  row: XlsxRow,
): { row: XlsxRow; sheetKey: string } {
  return { row, sheetKey: sheet.key };
}

export function xlsxColumnTargetState(
  sheet: XlsxRoundTripSheet,
  column: XlsxColumnRange,
): { column: XlsxColumnRange; sheetKey: string } {
  return { column, sheetKey: sheet.key };
}

function applyRowOperation(
  row: XlsxRow,
  operation: Extract<XlsxCellEditOperation, { kind: 'set-row' }>,
): void {
  if (operation.height !== undefined) row.height = operation.height;
  if (operation.hidden !== undefined) row.hidden = operation.hidden;
}

function applyColumnOperation(
  column: XlsxColumnRange,
  operation: Extract<XlsxCellEditOperation, { kind: 'set-column' }>,
): void {
  if (operation.width !== undefined) column.width = operation.width;
  if (operation.hidden !== undefined) column.hidden = operation.hidden;
}

function applyCellOperation(
  document: XlsxRoundTripDocument,
  sheet: XlsxWorksheet,
  cell: XlsxCell,
  operation: Extract<XlsxCellEditOperation, { cell: string }>,
): void {
  if (operation.kind === 'set-hyperlink') {
    const reference = parseXlsxCellReference(operation.cell)!;
    const conflict = sheet.hyperlinks.find(
      (candidate) =>
        candidate.range.reference !== operation.cell &&
        reference.row >= candidate.range.start.row &&
        reference.row <= candidate.range.end.row &&
        reference.column >= candidate.range.start.column &&
        reference.column <= candidate.range.end.column,
    );
    if (conflict) {
      operationFailure(
        'preservation-conflict',
        'XLSX hyperlink operation overlaps a multi-cell hyperlink range',
        operation,
        'hyperlink-range',
      );
    }
    const index = sheet.hyperlinks.findIndex(
      (candidate) => candidate.range.reference === operation.cell,
    );
    if (operation.target === null) {
      if (index >= 0) sheet.hyperlinks.splice(index, 1);
      return;
    }
    if (index >= 0) {
      sheet.hyperlinks[index] = {
        ...sheet.hyperlinks[index]!,
        target: structuredClone(operation.target),
      };
      return;
    }
    sheet.hyperlinks.push({
      range: {
        end: { column: reference.column, row: reference.row },
        reference: reference.address,
        start: { column: reference.column, row: reference.row },
      },
      selectionRelation: 'full-sheet',
      target: structuredClone(operation.target),
    });
    return;
  }
  delete cell.displayText;
  if (operation.kind === 'set-cell-style') {
    const styleKey = canonicalXlsxJson(operation.style);
    let style = document.styles.findIndex(
      (candidate) => canonicalXlsxJson(candidate) === styleKey,
    );
    if (style < 0) {
      if (operation.style.checkbox === true) {
        operationFailure(
          'unsupported-edit-operation',
          'XLSX cannot append a checkbox style without a feature-property-bag edit',
          operation,
          'append-checkbox-style',
        );
      }
      document.styles.push(structuredClone(operation.style));
      style = document.styles.length - 1;
    }
    cell.style = style;
    return;
  }
  if (operation.kind === 'clear-cell') {
    cell.content = { kind: 'blank' };
    return;
  }
  cell.content =
    operation.content.kind === 'formula'
      ? {
          cached: { kind: 'missing' },
          formula: {
            expression: operation.content.expression,
            kind: 'normal',
          },
          kind: 'formula',
        }
      : { kind: 'value', value: structuredClone(operation.content.value) };
}

export async function replayXlsxCellOperations(
  baseDocument: XlsxRoundTripDocument,
  value: unknown,
  writeLimits: ResolvedXlsxWriteLimits,
  readerLimits: ResolvedXlsxResourceLimits,
): Promise<XlsxCellOperationPlan> {
  const operations = validateXlsxCellOperations(
    value,
    writeLimits,
    readerLimits,
  );
  const document = cloneDocument(baseDocument);
  const sheetKeys = new Set<string>();
  for (const sheet of document.sheets) {
    if (sheetKeys.has(sheet.key)) {
      throw new XlsxWriteError(
        'snapshot-integrity-failed',
        'XLSX snapshot sheet keys must be unique',
        { objectKey: sheet.key },
      );
    }
    sheetKeys.add(sheet.key);
  }
  const impacts: XlsxCellOperationImpact[] = [];
  for (const operation of operations) {
    const sheet = resolveWorksheet(document, operation);
    if (operation.kind === 'set-row') {
      const row = resolveRow(sheet, operation);
      if (
        operation.ifMatch !== undefined &&
        operation.ifMatch !==
          (await canonicalXlsxSha256(xlsxRowTargetState(sheet, row)))
      ) {
        operationFailure(
          'operation-precondition-failed',
          'XLSX operation precondition does not match the target row',
          operation,
        );
      }
      applyRowOperation(row, operation);
      impacts.push({
        kind: operation.kind,
        operationId: operation.operationId,
        range: String(operation.row),
        sheetKey: operation.sheetKey,
      });
      continue;
    }
    if (operation.kind === 'set-column') {
      const column = resolveColumn(sheet, operation);
      if (
        operation.ifMatch !== undefined &&
        operation.ifMatch !==
          (await canonicalXlsxSha256(xlsxColumnTargetState(sheet, column)))
      ) {
        operationFailure(
          'operation-precondition-failed',
          'XLSX operation precondition does not match the target column range',
          operation,
        );
      }
      applyColumnOperation(column, operation);
      impacts.push({
        kind: operation.kind,
        operationId: operation.operationId,
        range: `${operation.start}:${operation.end}`,
        sheetKey: operation.sheetKey,
      });
      continue;
    }
    const cell = resolveCell(sheet, operation);
    if (
      operation.ifMatch !== undefined &&
      operation.ifMatch !==
        (await canonicalXlsxSha256(xlsxCellTargetState(sheet, cell)))
    ) {
      operationFailure(
        'operation-precondition-failed',
        'XLSX operation precondition does not match the target cell',
        operation,
      );
    }
    applyCellOperation(document, sheet, cell, operation);
    if (document.styles.length > readerLimits.maxStyles) {
      throw new XlsxWriteError(
        'resource-limit-exceeded',
        'XLSX edited normalized styles exceed their reader limit',
        {
          actual: document.styles.length,
          limit: readerLimits.maxStyles,
          limitName: 'maxStyles',
          operationId: operation.operationId,
        },
      );
    }
    impacts.push({
      cell: operation.cell,
      kind: operation.kind,
      operationId: operation.operationId,
      sheetKey: operation.sheetKey,
    });
  }
  return {
    document,
    impacts,
    operations,
    stateHash: await canonicalXlsxSha256(document),
  };
}

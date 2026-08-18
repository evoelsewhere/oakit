import type { PptxSceneValidationIssue } from './scene-types';
import type { PptxTableValidationDependencies } from './scene-table-validation';

function normalizedSpan(value: unknown): number {
  return Math.max(1, Number.isSafeInteger(value) ? Number(value) : 1);
}

export function validatePptxSceneTableMerges(
  rows: unknown[],
  columns: unknown[],
  path: string,
  issues: PptxSceneValidationIssue[],
  dependencies: PptxTableValidationDependencies,
): void {
  const expectedHorizontal = rows.map(() => columns.map(() => false));
  const expectedVertical = rows.map(() => columns.map(() => false));
  const occupied = rows.map(() => columns.map(() => false));
  rows.forEach((rowValue, rowIndex) => {
    if (!dependencies.isObject(rowValue) || !Array.isArray(rowValue.cells)) {
      return;
    }
    rowValue.cells.forEach((cellValue, columnIndex) => {
      if (!dependencies.isObject(cellValue)) return;
      const rowSpan = normalizedSpan(cellValue.rowSpan);
      const colSpan = normalizedSpan(cellValue.colSpan);
      if (rowSpan === 1 && colSpan === 1) return;
      if (
        rowIndex + rowSpan > rows.length ||
        columnIndex + colSpan > columns.length
      ) {
        dependencies.addIssue(
          issues,
          'invalid-scene-document',
          `${path}.rows[${rowIndex}].cells[${columnIndex}]`,
          'Table span exceeds the grid bounds',
        );
        return;
      }
      const rowOffsets = Array.from({ length: rowSpan }, (_, index) => index);
      const columnOffsets = Array.from(
        { length: colSpan },
        (_, index) => index,
      );
      for (const rowOffset of rowOffsets) {
        for (const columnOffset of columnOffsets) {
          const targetRow = rowIndex + rowOffset;
          const targetColumn = columnIndex + columnOffset;
          if (occupied[targetRow]?.[targetColumn]) {
            dependencies.addIssue(
              issues,
              'invalid-scene-document',
              `${path}.rows[${rowIndex}].cells[${columnIndex}]`,
              'Table spans must not overlap',
            );
            continue;
          }
          occupied[targetRow]![targetColumn] = true;
          expectedHorizontal[targetRow]![targetColumn] = columnOffset > 0;
          expectedVertical[targetRow]![targetColumn] = rowOffset > 0;
        }
      }
    });
  });
  rows.forEach((rowValue, rowIndex) => {
    if (!dependencies.isObject(rowValue) || !Array.isArray(rowValue.cells)) {
      return;
    }
    rowValue.cells.forEach((cellValue, columnIndex) => {
      if (!dependencies.isObject(cellValue)) return;
      if (
        (cellValue.hMerge === true) !==
          expectedHorizontal[rowIndex]?.[columnIndex] ||
        (cellValue.vMerge === true) !==
          expectedVertical[rowIndex]?.[columnIndex]
      ) {
        dependencies.addIssue(
          issues,
          'invalid-scene-document',
          `${path}.rows[${rowIndex}].cells[${columnIndex}]`,
          'Table merge continuation flags do not match its spans',
        );
      }
    });
  });
}

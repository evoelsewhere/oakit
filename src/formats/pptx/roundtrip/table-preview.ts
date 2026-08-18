import type { PptxSceneTableElement, PptxSceneTransform } from '../scene-types';
import type { Table } from '../types';
import {
  createPptxRoundTripTableCellPreview,
  type PptxPlainTextReader,
} from './table-preview-cell';

type TableTransformResolver = (
  element: Table,
) => PptxSceneTransform | undefined;

function hasNativeTableGrid(
  element: Table,
  transform: PptxSceneTransform | undefined,
): transform is PptxSceneTransform {
  if (transform === undefined) return false;
  if (
    element.colWidths.some((value) => value <= 0) ||
    element.rowHeights.length !== element.data.length ||
    element.rowHeights.some((value) => value <= 0) ||
    element.data.some((row) => row.length !== element.colWidths.length)
  ) {
    return false;
  }
  const columnWidth = element.colWidths.reduce(
    (total, value) => total + value,
    0,
  );
  const rowHeight = element.rowHeights.reduce(
    (total, value) => total + value,
    0,
  );
  return (
    Math.round(columnWidth * 12_700) === Math.round(transform.width * 12_700) &&
    Math.round(rowHeight * 12_700) === Math.round(transform.height * 12_700)
  );
}

export function createPptxRoundTripTablePreview(
  element: Table,
  slideIndex: number,
  elementIndex: number,
  plainText: PptxPlainTextReader,
  resolveTransform: TableTransformResolver,
): PptxSceneTableElement | undefined {
  const transform = resolveTransform(element);
  if (!hasNativeTableGrid(element, transform)) return undefined;
  const key = `slide-${slideIndex + 1}-element-${elementIndex + 1}`;
  return {
    authored: {},
    columns: [...element.colWidths],
    key,
    ...(element.name === undefined ? {} : { name: element.name }),
    resolved: { hidden: false, transform },
    rows: element.data.map((row, rowIndex) => ({
      cells: row.map((cell, columnIndex) =>
        createPptxRoundTripTableCellPreview(
          cell,
          key,
          rowIndex,
          columnIndex,
          plainText,
        ),
      ),
      height: element.rowHeights[rowIndex] as number,
    })),
    type: 'table',
  };
}

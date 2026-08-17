import { renderPptxSvgCellBorders } from './render-svg-table-border';
import {
  boundedTableSpan,
  normalizedTableSizes,
  tableOffsets,
} from './render-svg-table-layout';
import { renderPptxSvgTableText } from './render-svg-table-text';
import type { Table } from './types';
import { svgColor, svgNumber, type PptxSvgBox } from './render-svg-values';

function isRecord(value: unknown): value is Record<string, unknown> {
  return Object(value) === value && !Array.isArray(value);
}

function placeholder(box: PptxSvgBox): string {
  return `<rect x="0" y="0" width="${svgNumber(box.width)}" height="${svgNumber(box.height)}" fill="#f3f4f6" stroke="#9ca3af" stroke-dasharray="4 3"/><text x="4" y="16" font-family="sans-serif" font-size="12" fill="#374151">Empty table</text>`;
}

export function renderPptxSvgTable(table: Table, box: PptxSvgBox): string {
  if (!Array.isArray(table.data)) return placeholder(box);
  const rows = table.data;
  const columnCount = rows.reduce(
    (maximum, row) => Math.max(maximum, Array.isArray(row) ? row.length : 0),
    0,
  );
  if (columnCount === 0) return placeholder(box);
  const columnSizes = normalizedTableSizes(
    table.colWidths,
    columnCount,
    box.width,
  );
  const rowSizes = normalizedTableSizes(
    table.rowHeights,
    rows.length,
    box.height,
  );
  const columnOffsets = tableOffsets(columnSizes);
  const rowOffsets = tableOffsets(rowSizes);
  return rows
    .map((row, rowIndex) => {
      if (!Array.isArray(row)) return '';
      return row
        .map((cell, columnIndex) => {
          const value: Record<string, unknown> = isRecord(cell) ? cell : {};
          if (value.hMerge === 1 || value.vMerge === 1) return '';
          const columnSpan = boundedTableSpan(
            value.colSpan,
            columnCount - columnIndex,
          );
          const rowSpan = boundedTableSpan(
            value.rowSpan,
            rows.length - rowIndex,
          );
          const x = columnOffsets[columnIndex] as number;
          const y = rowOffsets[rowIndex] as number;
          const width = (columnOffsets[columnIndex + columnSpan] as number) - x;
          const height = (rowOffsets[rowIndex + rowSpan] as number) - y;
          const fill = svgColor(value.fillColor) ?? '#ffffff';
          return `<g transform="translate(${svgNumber(x)} ${svgNumber(y)})"><rect x="0" y="0" width="${svgNumber(width)}" height="${svgNumber(height)}" fill="${fill}" stroke="none"/>${renderPptxSvgCellBorders(value.borders, 0, 0, width, height)}${renderPptxSvgTableText(value, width, height)}</g>`;
        })
        .join('');
    })
    .join('');
}

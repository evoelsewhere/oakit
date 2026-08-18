import type { PptxSceneTableBorder, PptxSceneTableCell } from '../scene-types';
import type { Border, TableCell } from '../types';

export type PptxPlainTextReader = (html: string) => string;

function sceneTableBorder(border: Border): PptxSceneTableBorder {
  return {
    color: border.borderColor,
    style: border.borderType,
    width: border.borderWidth,
  };
}

function hasVisibleTableBorder(border: Border | undefined): border is Border {
  return border !== undefined && border.borderWidth > 0;
}

export function createPptxRoundTripTableCellPreview(
  cell: TableCell,
  tableKey: string,
  rowIndex: number,
  columnIndex: number,
  plainText: PptxPlainTextReader,
): PptxSceneTableCell {
  const key = `${tableKey}-row-${rowIndex + 1}-cell-${columnIndex + 1}`;
  const properties = {
    ...(cell.fontBold === undefined ? {} : { bold: cell.fontBold }),
    ...(cell.fontColor === undefined ? {} : { color: cell.fontColor }),
  };
  return {
    borders: {
      ...(!hasVisibleTableBorder(cell.borders.bottom)
        ? {}
        : { bottom: sceneTableBorder(cell.borders.bottom) }),
      ...(!hasVisibleTableBorder(cell.borders.left)
        ? {}
        : { left: sceneTableBorder(cell.borders.left) }),
      ...(!hasVisibleTableBorder(cell.borders.right)
        ? {}
        : { right: sceneTableBorder(cell.borders.right) }),
      ...(!hasVisibleTableBorder(cell.borders.top)
        ? {}
        : { top: sceneTableBorder(cell.borders.top) }),
    },
    ...(cell.colSpan === undefined ? {} : { colSpan: cell.colSpan }),
    ...(cell.fillColor === undefined ? {} : { fillColor: cell.fillColor }),
    ...(cell.hMerge === undefined ? {} : { hMerge: true }),
    ...(cell.rowSpan === undefined ? {} : { rowSpan: cell.rowSpan }),
    text: {
      body: {
        anchor:
          cell.vAlign === 'down'
            ? 'bottom'
            : cell.vAlign === 'mid'
              ? 'center'
              : 'top',
      },
      paragraphs: [
        {
          children: [
            {
              key: `${key}-run-1`,
              ...(Object.keys(properties).length === 0 ? {} : { properties }),
              text: plainText(cell.text),
              type: 'run',
            },
          ],
          key: `${key}-paragraph-1`,
        },
      ],
    },
    ...(cell.vMerge === undefined ? {} : { vMerge: true }),
  };
}

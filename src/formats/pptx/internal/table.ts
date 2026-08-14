import type { XmlLookupValue } from '../../../common';
import type { PptxParserContext } from './context';

import { getShapeFill, getSolidFill } from './fill';
import { getTextByPathList } from '../../../common';
import { getBorder } from './border';

type ParsedBorder = ReturnType<typeof getBorder>;
type TableBorders = Partial<
  Record<'bottom' | 'left' | 'right' | 'top', ParsedBorder>
>;

interface TableStyleAttributes {
  isBandColAttr: number;
  isBandRowAttr: number;
  isFrstColAttr: number;
  isFrstRowAttr: number;
  isLstColAttr: number;
  isLstRowAttr: number;
}

interface TableCellParams {
  borders: TableBorders;
  colSpan?: number;
  fillColor?: string;
  fontBold?: boolean;
  fontColor?: string;
  hMerge?: number;
  rowSpan?: number;
  vAlign: 'down' | 'mid' | 'up';
  vMerge?: number;
}

function getTableTextColor(
  tcTxStyle: XmlLookupValue | undefined,
  warpObj: PptxParserContext,
): string | undefined {
  if (!tcTxStyle) return undefined;

  return getSolidFill(
    tcTxStyle['a:solidFill'] || tcTxStyle,
    undefined,
    undefined,
    warpObj,
  );
}

export function getTableBorders(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
): TableBorders {
  const borders: TableBorders = {};
  if (node['a:bottom']) {
    const obj = {
      'p:spPr': {
        'a:ln': node['a:bottom']['a:ln'],
      },
    };
    const border = getBorder(obj, undefined, warpObj);
    borders.bottom = border;
  }
  if (node['a:top']) {
    const obj = {
      'p:spPr': {
        'a:ln': node['a:top']['a:ln'],
      },
    };
    const border = getBorder(obj, undefined, warpObj);
    borders.top = border;
  }
  if (node['a:right']) {
    const obj = {
      'p:spPr': {
        'a:ln': node['a:right']['a:ln'],
      },
    };
    const border = getBorder(obj, undefined, warpObj);
    borders.right = border;
  }
  if (node['a:left']) {
    const obj = {
      'p:spPr': {
        'a:ln': node['a:left']['a:ln'],
      },
    };
    const border = getBorder(obj, undefined, warpObj);
    borders.left = border;
  }
  return borders;
}

export async function getTableCellParams(
  tcNode: XmlLookupValue,
  thisTblStyle: XmlLookupValue | undefined,
  cellSource: string | undefined,
  warpObj: PptxParserContext,
): Promise<TableCellParams> {
  const rowSpan = getTextByPathList<string>(tcNode, ['attrs', 'rowSpan']);
  const colSpan = getTextByPathList<string>(tcNode, ['attrs', 'gridSpan']);
  const vMerge = getTextByPathList<string>(tcNode, ['attrs', 'vMerge']);
  const hMerge = getTextByPathList<string>(tcNode, ['attrs', 'hMerge']);
  const anchor = getTextByPathList<string>(tcNode, [
    'a:tcPr',
    'attrs',
    'anchor',
  ]);
  let fillColor: string | undefined;
  let fontColor: string | undefined;
  let fontBold: boolean | undefined;

  const getCelFill = getTextByPathList<XmlLookupValue>(tcNode, ['a:tcPr']);
  if (getCelFill) {
    const cellObj = { 'p:spPr': getCelFill };
    const fill = await getShapeFill(
      cellObj as unknown as XmlLookupValue,
      warpObj,
      'slide',
    );

    if (fill?.type === 'color' && typeof fill.value === 'string') {
      fillColor = fill.value;
    }
  }
  if (!fillColor) {
    let bgFillschemeClr: XmlLookupValue | undefined;
    if (cellSource)
      bgFillschemeClr = getTextByPathList<XmlLookupValue>(thisTblStyle, [
        cellSource,
        'a:tcStyle',
        'a:fill',
        'a:solidFill',
      ]);
    if (bgFillschemeClr) {
      fillColor = getSolidFill(bgFillschemeClr, undefined, undefined, warpObj);
    }
  }

  let rowTxtStyl: XmlLookupValue | undefined;
  if (cellSource)
    rowTxtStyl = getTextByPathList<XmlLookupValue>(thisTblStyle, [
      cellSource,
      'a:tcTxStyle',
    ]);
  if (rowTxtStyl) {
    fontColor = getTableTextColor(rowTxtStyl, warpObj);
    if (getTextByPathList(rowTxtStyl, ['attrs', 'b']) === 'on') fontBold = true;
  }

  let lin_bottm = getTextByPathList<XmlLookupValue>(tcNode, [
    'a:tcPr',
    'a:lnB',
  ]);
  if (!lin_bottm) {
    if (cellSource)
      lin_bottm = getTextByPathList<XmlLookupValue>(
        thisTblStyle?.[cellSource],
        ['a:tcStyle', 'a:tcBdr', 'a:bottom', 'a:ln'],
      );
    if (!lin_bottm)
      lin_bottm = getTextByPathList<XmlLookupValue>(thisTblStyle, [
        'a:wholeTbl',
        'a:tcStyle',
        'a:tcBdr',
        'a:bottom',
        'a:ln',
      ]);
  }
  let lin_top = getTextByPathList<XmlLookupValue>(tcNode, ['a:tcPr', 'a:lnT']);
  if (!lin_top) {
    if (cellSource)
      lin_top = getTextByPathList<XmlLookupValue>(thisTblStyle?.[cellSource], [
        'a:tcStyle',
        'a:tcBdr',
        'a:top',
        'a:ln',
      ]);
    if (!lin_top)
      lin_top = getTextByPathList<XmlLookupValue>(thisTblStyle, [
        'a:wholeTbl',
        'a:tcStyle',
        'a:tcBdr',
        'a:top',
        'a:ln',
      ]);
  }
  let lin_left = getTextByPathList<XmlLookupValue>(tcNode, ['a:tcPr', 'a:lnL']);
  if (!lin_left) {
    if (cellSource)
      lin_left = getTextByPathList<XmlLookupValue>(thisTblStyle?.[cellSource], [
        'a:tcStyle',
        'a:tcBdr',
        'a:left',
        'a:ln',
      ]);
    if (!lin_left)
      lin_left = getTextByPathList<XmlLookupValue>(thisTblStyle, [
        'a:wholeTbl',
        'a:tcStyle',
        'a:tcBdr',
        'a:left',
        'a:ln',
      ]);
  }
  let lin_right = getTextByPathList<XmlLookupValue>(tcNode, [
    'a:tcPr',
    'a:lnR',
  ]);
  if (!lin_right) {
    if (cellSource)
      lin_right = getTextByPathList<XmlLookupValue>(
        thisTblStyle?.[cellSource],
        ['a:tcStyle', 'a:tcBdr', 'a:right', 'a:ln'],
      );
    if (!lin_right)
      lin_right = getTextByPathList<XmlLookupValue>(thisTblStyle, [
        'a:wholeTbl',
        'a:tcStyle',
        'a:tcBdr',
        'a:right',
        'a:ln',
      ]);
  }

  const borders: TableBorders = {};
  if (lin_bottm) borders.bottom = getBorder(lin_bottm, undefined, warpObj);
  if (lin_top) borders.top = getBorder(lin_top, undefined, warpObj);
  if (lin_left) borders.left = getBorder(lin_left, undefined, warpObj);
  if (lin_right) borders.right = getBorder(lin_right, undefined, warpObj);

  return {
    borders,
    vAlign: anchor === 'ctr' ? 'mid' : anchor === 'b' ? 'down' : 'up',
    ...(fillColor ? { fillColor } : {}),
    ...(fontColor ? { fontColor } : {}),
    ...(fontBold !== undefined ? { fontBold } : {}),
    ...(rowSpan ? { rowSpan: Number(rowSpan) } : {}),
    ...(colSpan ? { colSpan: Number(colSpan) } : {}),
    ...(vMerge ? { vMerge: Number(vMerge) } : {}),
    ...(hMerge ? { hMerge: Number(hMerge) } : {}),
  };
}

export function getTableRowParams(
  trNodes: XmlLookupValue[],
  i: number,
  tblStylAttrObj: TableStyleAttributes,
  thisTblStyle: XmlLookupValue | undefined,
  warpObj: PptxParserContext,
): { fillColor?: string; fontBold?: boolean; fontColor?: string } {
  let fillColor: string | undefined;
  let fontColor: string | undefined;
  let fontBold: boolean | undefined;

  if (thisTblStyle && thisTblStyle['a:wholeTbl']) {
    const bgFillschemeClr = getTextByPathList(thisTblStyle, [
      'a:wholeTbl',
      'a:tcStyle',
      'a:fill',
      'a:solidFill',
    ]);
    if (bgFillschemeClr) {
      const local_fillColor = getSolidFill(
        bgFillschemeClr,
        undefined,
        undefined,
        warpObj,
      );
      if (local_fillColor) fillColor = local_fillColor;
    }
    const rowTxtStyl = getTextByPathList(thisTblStyle, [
      'a:wholeTbl',
      'a:tcTxStyle',
    ]);
    if (rowTxtStyl) {
      const local_fontColor = getTableTextColor(rowTxtStyl, warpObj);
      if (local_fontColor) fontColor = local_fontColor;
      if (getTextByPathList(rowTxtStyl, ['attrs', 'b']) === 'on')
        fontBold = true;
    }
  }
  if (i === 0 && tblStylAttrObj['isFrstRowAttr'] === 1 && thisTblStyle) {
    const bgFillschemeClr = getTextByPathList(thisTblStyle, [
      'a:firstRow',
      'a:tcStyle',
      'a:fill',
      'a:solidFill',
    ]);
    if (bgFillschemeClr) {
      const local_fillColor = getSolidFill(
        bgFillschemeClr,
        undefined,
        undefined,
        warpObj,
      );
      if (local_fillColor) fillColor = local_fillColor;
    }
    const rowTxtStyl = getTextByPathList(thisTblStyle, [
      'a:firstRow',
      'a:tcTxStyle',
    ]);
    if (rowTxtStyl) {
      const local_fontColor = getTableTextColor(rowTxtStyl, warpObj);
      if (local_fontColor) fontColor = local_fontColor;
      if (getTextByPathList(rowTxtStyl, ['attrs', 'b']) === 'on')
        fontBold = true;
    }
  } else if (i > 0 && tblStylAttrObj['isBandRowAttr'] === 1 && thisTblStyle) {
    fillColor = '';
    if (i % 2 === 0 && thisTblStyle['a:band2H']) {
      const bgFillschemeClr = getTextByPathList(thisTblStyle, [
        'a:band2H',
        'a:tcStyle',
        'a:fill',
        'a:solidFill',
      ]);
      if (bgFillschemeClr) {
        const local_fillColor = getSolidFill(
          bgFillschemeClr,
          undefined,
          undefined,
          warpObj,
        );
        if (local_fillColor) fillColor = local_fillColor;
      }
      const rowTxtStyl = getTextByPathList(thisTblStyle, [
        'a:band2H',
        'a:tcTxStyle',
      ]);
      if (rowTxtStyl) {
        const local_fontColor = getTableTextColor(rowTxtStyl, warpObj);
        if (local_fontColor) fontColor = local_fontColor;
      }
      if (getTextByPathList(rowTxtStyl, ['attrs', 'b']) === 'on')
        fontBold = true;
    }
    if (i % 2 !== 0 && thisTblStyle['a:band1H']) {
      const bgFillschemeClr = getTextByPathList(thisTblStyle, [
        'a:band1H',
        'a:tcStyle',
        'a:fill',
        'a:solidFill',
      ]);
      if (bgFillschemeClr) {
        const local_fillColor = getSolidFill(
          bgFillschemeClr,
          undefined,
          undefined,
          warpObj,
        );
        if (local_fillColor) fillColor = local_fillColor;
      }
      const rowTxtStyl = getTextByPathList(thisTblStyle, [
        'a:band1H',
        'a:tcTxStyle',
      ]);
      if (rowTxtStyl) {
        const local_fontColor = getTableTextColor(rowTxtStyl, warpObj);
        if (local_fontColor) fontColor = local_fontColor;
        if (getTextByPathList(rowTxtStyl, ['attrs', 'b']) === 'on')
          fontBold = true;
      }
    }
  }
  if (
    i === trNodes.length - 1 &&
    tblStylAttrObj['isLstRowAttr'] === 1 &&
    thisTblStyle
  ) {
    const bgFillschemeClr = getTextByPathList(thisTblStyle, [
      'a:lastRow',
      'a:tcStyle',
      'a:fill',
      'a:solidFill',
    ]);
    if (bgFillschemeClr) {
      const local_fillColor = getSolidFill(
        bgFillschemeClr,
        undefined,
        undefined,
        warpObj,
      );
      if (local_fillColor) {
        fillColor = local_fillColor;
      }
    }
    const rowTxtStyl = getTextByPathList(thisTblStyle, [
      'a:lastRow',
      'a:tcTxStyle',
    ]);
    if (rowTxtStyl) {
      const local_fontColor = getTableTextColor(rowTxtStyl, warpObj);
      if (local_fontColor) fontColor = local_fontColor;
      if (getTextByPathList(rowTxtStyl, ['attrs', 'b']) === 'on')
        fontBold = true;
    }
  }

  return {
    ...(fillColor !== undefined ? { fillColor } : {}),
    ...(fontColor !== undefined ? { fontColor } : {}),
    ...(fontBold !== undefined ? { fontBold } : {}),
  };
}

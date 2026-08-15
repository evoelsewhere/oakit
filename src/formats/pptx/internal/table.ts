import type { XmlLookupValue } from '../../../common';
import type { PptxParserContext } from './context';

import { getSolidFill } from './fill';
import { getTextByPathList } from '../../../common';
import { getBorder } from './border';

type ParsedBorder = ReturnType<typeof getBorder>;
type TableBorders = Partial<
  Record<'bottom' | 'left' | 'right' | 'top', ParsedBorder>
>;

type TableBorderDirection = keyof TableBorders;
type TableStyleSection =
  'a:band1H' | 'a:band2H' | 'a:firstRow' | 'a:lastRow' | 'a:wholeTbl';

interface TableBorderDescriptor {
  cellKey: 'a:lnB' | 'a:lnL' | 'a:lnR' | 'a:lnT';
  direction: TableBorderDirection;
  styleKey: 'a:bottom' | 'a:left' | 'a:right' | 'a:top';
}

interface TableStyleParams {
  fillColor: string;
  fontBold: boolean | undefined;
  fontColor: string;
}

const TABLE_BORDER_DESCRIPTORS: readonly TableBorderDescriptor[] = [
  { cellKey: 'a:lnB', direction: 'bottom', styleKey: 'a:bottom' },
  { cellKey: 'a:lnL', direction: 'left', styleKey: 'a:left' },
  { cellKey: 'a:lnR', direction: 'right', styleKey: 'a:right' },
  { cellKey: 'a:lnT', direction: 'top', styleKey: 'a:top' },
];

const TRUE_VALUES: ReadonlySet<string | undefined> = new Set([
  '1',
  'on',
  'true',
]);
const FALSE_VALUES: ReadonlySet<string | undefined> = new Set([
  '0',
  'off',
  'false',
]);

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

function positiveInteger(value: string | undefined): number | undefined {
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed > 0 ? parsed : undefined;
}

function activeMerge(value: string | undefined): 1 | undefined {
  return value === '1' ? 1 : undefined;
}

function getTableTextColor(
  tcTxStyle: XmlLookupValue | undefined,
  warpObj: PptxParserContext,
): string {
  return getSolidFill(
    tcTxStyle?.['a:solidFill'] ?? tcTxStyle,
    undefined,
    undefined,
    warpObj,
  );
}

function getOnOff(value: string | undefined): boolean | undefined {
  if (TRUE_VALUES.has(value)) return true;
  if (FALSE_VALUES.has(value)) return false;
  return undefined;
}

function getTableStyleParams(
  tableStyle: XmlLookupValue | undefined,
  sectionName: string,
  warpObj: PptxParserContext,
): TableStyleParams {
  const section = tableStyle?.[sectionName];
  const textStyle = getTextByPathList<XmlLookupValue>(section, ['a:tcTxStyle']);

  return {
    fillColor: getSolidFill(
      getTextByPathList<XmlLookupValue>(section, [
        'a:tcStyle',
        'a:fill',
        'a:solidFill',
      ]),
      undefined,
      undefined,
      warpObj,
    ),
    fontBold: getOnOff(getTextByPathList<string>(textStyle, ['attrs', 'b'])),
    fontColor: getTableTextColor(textStyle, warpObj),
  };
}

function mergeTableStyleParams(
  base: TableStyleParams,
  override: TableStyleParams,
): TableStyleParams {
  return {
    fillColor: override.fillColor || base.fillColor,
    fontBold: override.fontBold ?? base.fontBold,
    fontColor: override.fontColor || base.fontColor,
  };
}

function getCellBorderLine(
  tcNode: XmlLookupValue,
  tableStyle: XmlLookupValue | undefined,
  sourceStyle: XmlLookupValue | undefined,
  descriptor: TableBorderDescriptor,
): XmlLookupValue | undefined {
  return (
    getTextByPathList<XmlLookupValue>(tcNode, ['a:tcPr', descriptor.cellKey]) ??
    getTextByPathList<XmlLookupValue>(sourceStyle, [
      'a:tcStyle',
      'a:tcBdr',
      descriptor.styleKey,
      'a:ln',
    ]) ??
    getTextByPathList<XmlLookupValue>(tableStyle, [
      'a:wholeTbl',
      'a:tcStyle',
      'a:tcBdr',
      descriptor.styleKey,
      'a:ln',
    ])
  );
}

export function getTableBorders(
  node: XmlLookupValue,
  warpObj: PptxParserContext,
): TableBorders {
  const borders: TableBorders = {};
  for (const descriptor of TABLE_BORDER_DESCRIPTORS) {
    const styledBorder = node[descriptor.styleKey];
    if (styledBorder) {
      borders[descriptor.direction] = getBorder(
        { 'p:spPr': { 'a:ln': styledBorder['a:ln'] } },
        undefined,
        warpObj,
      );
    }
  }
  return borders;
}

export function getTableCellParams(
  tcNode: XmlLookupValue,
  thisTblStyle: XmlLookupValue | undefined,
  cellSource: string | undefined,
  warpObj: PptxParserContext,
): TableCellParams {
  const rowSpan = positiveInteger(
    getTextByPathList<string>(tcNode, ['attrs', 'rowSpan']),
  );
  const colSpan = positiveInteger(
    getTextByPathList<string>(tcNode, ['attrs', 'gridSpan']),
  );
  const vMerge = activeMerge(
    getTextByPathList<string>(tcNode, ['attrs', 'vMerge']),
  );
  const hMerge = activeMerge(
    getTextByPathList<string>(tcNode, ['attrs', 'hMerge']),
  );
  const anchor = getTextByPathList<string>(tcNode, [
    'a:tcPr',
    'attrs',
    'anchor',
  ]);
  const sourceStyle = thisTblStyle?.[cellSource as string];
  const sourceParams = getTableStyleParams(
    thisTblStyle,
    cellSource as string,
    warpObj,
  );
  const directFillColor = getSolidFill(
    getTextByPathList<XmlLookupValue>(tcNode, ['a:tcPr', 'a:solidFill']),
    undefined,
    undefined,
    warpObj,
  );
  const fillColor = directFillColor || sourceParams.fillColor;

  const borders: TableBorders = {};
  for (const descriptor of TABLE_BORDER_DESCRIPTORS) {
    const line = getCellBorderLine(
      tcNode,
      thisTblStyle,
      sourceStyle,
      descriptor,
    );
    if (line) {
      borders[descriptor.direction] = getBorder(line, undefined, warpObj);
    }
  }

  return {
    borders,
    vAlign: anchor === 'ctr' ? 'mid' : anchor === 'b' ? 'down' : 'up',
    ...(fillColor ? { fillColor } : {}),
    ...(sourceParams.fontColor ? { fontColor: sourceParams.fontColor } : {}),
    ...(sourceParams.fontBold !== undefined
      ? { fontBold: sourceParams.fontBold }
      : {}),
    ...(rowSpan !== undefined ? { rowSpan } : {}),
    ...(colSpan !== undefined ? { colSpan } : {}),
    ...(vMerge !== undefined ? { vMerge } : {}),
    ...(hMerge !== undefined ? { hMerge } : {}),
  };
}

export function getTableRowParams(
  trNodes: XmlLookupValue[],
  i: number,
  tblStylAttrObj: TableStyleAttributes,
  thisTblStyle: XmlLookupValue | undefined,
  warpObj: PptxParserContext,
): { fillColor?: string; fontBold?: boolean; fontColor?: string } {
  let params = getTableStyleParams(thisTblStyle, 'a:wholeTbl', warpObj);

  if (i > 0 && tblStylAttrObj.isBandRowAttr === 1) {
    const bandSection: TableStyleSection =
      i % 2 === 0 ? 'a:band2H' : 'a:band1H';
    params = mergeTableStyleParams(
      params,
      getTableStyleParams(thisTblStyle, bandSection, warpObj),
    );
  }
  if (i === 0 && tblStylAttrObj.isFrstRowAttr === 1) {
    params = mergeTableStyleParams(
      params,
      getTableStyleParams(thisTblStyle, 'a:firstRow', warpObj),
    );
  }
  if (i === trNodes.length - 1 && tblStylAttrObj.isLstRowAttr === 1) {
    params = mergeTableStyleParams(
      params,
      getTableStyleParams(thisTblStyle, 'a:lastRow', warpObj),
    );
  }

  return {
    ...(params.fillColor ? { fillColor: params.fillColor } : {}),
    ...(params.fontColor ? { fontColor: params.fontColor } : {}),
    ...(params.fontBold !== undefined ? { fontBold: params.fontBold } : {}),
  };
}

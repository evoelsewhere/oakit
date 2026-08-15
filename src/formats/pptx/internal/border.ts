import type { XmlLookupValue } from '../../../common';
import type { PptxParserContext } from './context';
import type { Border, LineEnd } from '../types';

import tinycolor from 'tinycolor2';
import { getSolidFill } from './fill';
import { getSchemeColorFromTheme } from './scheme-color';
import { getTextByPathList } from '../../../common';
import { normalizeHexColor } from '../../../common/text/css';

const DEFAULT_BORDER_COLOR = '#000000';
const LINE_END_TYPES: ReadonlySet<string | undefined> = new Set([
  'triangle',
  'stealth',
  'diamond',
  'oval',
  'arrow',
]);
const LINE_END_SIZES: ReadonlySet<string | undefined> = new Set([
  'sm',
  'med',
  'lg',
]);

function getLineEnd(node?: XmlLookupValue): LineEnd | undefined {
  const attrs = getTextByPathList<Record<string, string>>(node, ['attrs']);
  if (!attrs) return undefined;

  const lineEnd: LineEnd = {
    type: LINE_END_TYPES.has(attrs.type)
      ? (attrs.type as LineEnd['type'])
      : 'none',
  };
  if (LINE_END_SIZES.has(attrs.w)) {
    lineEnd.width = attrs.w as NonNullable<LineEnd['width']>;
  }
  if (LINE_END_SIZES.has(attrs.len)) {
    lineEnd.length = attrs.len as NonNullable<LineEnd['length']>;
  }
  return lineEnd;
}

function getReferencedLine(
  node: unknown,
  warpObj: PptxParserContext,
): XmlLookupValue | undefined {
  const lnRefNode = getTextByPathList<XmlLookupValue>(node, [
    'p:style',
    'a:lnRef',
  ]);
  const lineIndex =
    Number(getTextByPathList<string>(lnRefNode, ['attrs', 'idx'])) - 1;
  const themeLines = getTextByPathList<XmlLookupValue | XmlLookupValue[]>(
    warpObj.themeContent,
    ['a:theme', 'a:themeElements', 'a:fmtScheme', 'a:lnStyleLst', 'a:ln'],
  );

  if (Array.isArray(themeLines)) return themeLines[lineIndex];
  return lineIndex === 0 ? themeLines : undefined;
}

export function getBorder(
  node: unknown,
  _elType: string | undefined,
  warpObj: PptxParserContext,
): Border & {
  headEnd?: LineEnd;
  strokeDasharray: string;
  tailEnd?: LineEnd;
} {
  const lineNode =
    getTextByPathList<XmlLookupValue>(node, ['p:spPr', 'a:ln']) ??
    getReferencedLine(node, warpObj) ??
    (node as XmlLookupValue | undefined);

  const isNoFill = getTextByPathList(lineNode, ['a:noFill']);

  const widthEmus = Number(
    getTextByPathList<string>(lineNode, ['attrs', 'w']) ?? 0,
  );
  const borderWidth =
    !isNoFill && Number.isSafeInteger(widthEmus)
      ? Math.max(widthEmus, 0) / 12_700
      : 0;

  const solidFill = getTextByPathList<XmlLookupValue>(lineNode, [
    'a:solidFill',
  ]);
  let borderColor: string | undefined = getSolidFill(
    solidFill,
    undefined,
    undefined,
    warpObj,
  );

  if (!borderColor) {
    const schemeClrNode = getTextByPathList<XmlLookupValue>(node, [
      'p:style',
      'a:lnRef',
      'a:schemeClr',
    ]);
    const schemeName = getTextByPathList<string>(schemeClrNode, [
      'attrs',
      'val',
    ]);
    const themeColor = getSchemeColorFromTheme(`a:${schemeName}`, warpObj);
    borderColor = themeColor === undefined ? DEFAULT_BORDER_COLOR : themeColor;

    const shadeValue = getTextByPathList<string>(schemeClrNode, [
      'a:shade',
      'attrs',
      'val',
    ]);

    const shadeEmus = Number(shadeValue);
    if (Number.isSafeInteger(shadeEmus)) {
      if (shadeEmus >= 0) {
        const shade = Math.min(shadeEmus, 100_000) / 100_000;

        const color = tinycolor(borderColor).toHsl();
        borderColor = tinycolor({
          h: color.h,
          s: color.s,
          l: color.l * shade,
          a: color.a,
        }).toHex();
      }
    }
  }

  const safeBorderColor =
    normalizeHexColor(borderColor) ?? DEFAULT_BORDER_COLOR;

  const type = getTextByPathList<string>(lineNode, [
    'a:prstDash',
    'attrs',
    'val',
  ]);
  let borderType: Border['borderType'] = 'solid';
  let strokeDasharray = '0';
  switch (type) {
    case 'dash':
      borderType = 'dashed';
      strokeDasharray = '5';
      break;
    case 'dashDot':
      borderType = 'dashed';
      strokeDasharray = '5, 5, 1, 5';
      break;
    case 'dot':
      borderType = 'dotted';
      strokeDasharray = '1, 5';
      break;
    case 'lgDash':
      borderType = 'dashed';
      strokeDasharray = '10, 5';
      break;
    case 'lgDashDotDot':
      borderType = 'dotted';
      strokeDasharray = '10, 5, 1, 5, 1, 5';
      break;
    case 'sysDash':
      borderType = 'dashed';
      strokeDasharray = '5, 2';
      break;
    case 'sysDashDot':
      borderType = 'dotted';
      strokeDasharray = '5, 2, 1, 5';
      break;
    case 'sysDashDotDot':
      borderType = 'dotted';
      strokeDasharray = '5, 2, 1, 5, 1, 5';
      break;
    case 'sysDot':
      borderType = 'dotted';
      strokeDasharray = '2, 5';
      break;
    default:
  }

  const headEnd = getLineEnd(
    getTextByPathList<XmlLookupValue>(lineNode, ['a:headEnd']),
  );
  const tailEnd = getLineEnd(
    getTextByPathList<XmlLookupValue>(lineNode, ['a:tailEnd']),
  );

  return {
    borderColor: safeBorderColor,
    borderWidth,
    borderType,
    strokeDasharray,
    ...(headEnd ? { headEnd } : {}),
    ...(tailEnd ? { tailEnd } : {}),
  };
}

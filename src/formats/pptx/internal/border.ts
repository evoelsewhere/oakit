import type { XmlLookupValue } from '../../../common';
import type { PptxParserContext } from './context';
import type { Border, LineEnd } from '../types';

import tinycolor from 'tinycolor2';
import { getSolidFill } from './fill';
import { getSchemeColorFromTheme } from './scheme-color';
import { getTextByPathList } from '../../../common';
import { normalizeHexColor } from '../../../common/text/css';

function getLineEnd(node?: XmlLookupValue): LineEnd | undefined {
  const attrs = getTextByPathList<Record<string, string>>(node, ['attrs']);
  if (!attrs) return undefined;

  const lineEnd: LineEnd = {
    type: (attrs.type as LineEnd['type'] | undefined) ?? 'none',
  };
  if (attrs.w) lineEnd.width = attrs.w as NonNullable<LineEnd['width']>;
  if (attrs.len) lineEnd.length = attrs.len as NonNullable<LineEnd['length']>;
  return lineEnd;
}

export function getBorder(
  node: unknown,
  elType: string | undefined,
  warpObj: PptxParserContext,
): Border & {
  headEnd?: LineEnd;
  strokeDasharray: string;
  tailEnd?: LineEnd;
} {
  let lineNode = getTextByPathList<XmlLookupValue>(node, ['p:spPr', 'a:ln']);
  if (!lineNode) {
    const lnRefNode = getTextByPathList<XmlLookupValue>(node, [
      'p:style',
      'a:lnRef',
    ]);
    if (lnRefNode) {
      const lnIdx = getTextByPathList<string>(lnRefNode, ['attrs', 'idx']);
      const themeLines = getTextByPathList<XmlLookupValue | XmlLookupValue[]>(
        warpObj.themeContent,
        ['a:theme', 'a:themeElements', 'a:fmtScheme', 'a:lnStyleLst', 'a:ln'],
      );
      if (themeLines) {
        const lines = Array.isArray(themeLines) ? themeLines : [themeLines];
        lineNode = lines[Number(lnIdx) - 1];
      }
    }
  }
  if (!lineNode && typeof node === 'object' && node !== null) {
    lineNode = node as XmlLookupValue;
  }
  if (!lineNode) {
    return {
      borderColor: '#000000',
      borderWidth: 0,
      borderType: 'solid',
      strokeDasharray: '0',
    };
  }

  const isNoFill = getTextByPathList(lineNode, ['a:noFill']);

  let borderWidth = isNoFill
    ? 0
    : parseInt(getTextByPathList<string>(lineNode, ['attrs', 'w']) ?? '0') /
      12700;
  if (isNaN(borderWidth)) {
    if (lineNode) borderWidth = 0;
    else if (elType !== 'obj') borderWidth = 0;
    else borderWidth = 1;
  }

  const solidFill = getTextByPathList<XmlLookupValue>(lineNode, [
    'a:solidFill',
  ]);
  let borderColor = getSolidFill(solidFill, undefined, undefined, warpObj);

  if (!borderColor) {
    const schemeClrNode = getTextByPathList<XmlLookupValue>(node, [
      'p:style',
      'a:lnRef',
      'a:schemeClr',
    ]);
    const schemeClr = `a:${getTextByPathList<string>(schemeClrNode, ['attrs', 'val']) ?? ''}`;
    borderColor = getSchemeColorFromTheme(schemeClr, warpObj) ?? '';

    if (borderColor) {
      const shadeValue = getTextByPathList<string>(schemeClrNode, [
        'a:shade',
        'attrs',
        'val',
      ]);

      if (shadeValue) {
        const shade = parseInt(shadeValue) / 100000;

        const color = tinycolor('#' + borderColor).toHsl();
        borderColor = tinycolor({
          h: color.h,
          s: color.s,
          l: color.l * shade,
          a: color.a,
        }).toHex();
      }
    }
  }

  borderColor = normalizeHexColor(borderColor) ?? '#000000';

  const type = getTextByPathList<string>(lineNode, [
    'a:prstDash',
    'attrs',
    'val',
  ]);
  let borderType: Border['borderType'] = 'solid';
  let strokeDasharray = '0';
  switch (type) {
    case 'solid':
      borderType = 'solid';
      strokeDasharray = '0';
      break;
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
    borderColor,
    borderWidth,
    borderType,
    strokeDasharray,
    ...(headEnd ? { headEnd } : {}),
    ...(tailEnd ? { tailEnd } : {}),
  };
}

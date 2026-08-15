import type { XmlLookupValue } from '../../../common';
import type { PptxParserContext } from './context';

import { getTextByPathList } from '../../../common';
import { getShadow } from './shadow';
import { getGradientFill, getSolidFill } from './fill';

function parseInteger(value: string | undefined): number | undefined {
  if (value === undefined || !/^[+-]?\d+$/.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}

function isDrawingMlTrue(value: string | undefined): boolean {
  return value === '1' || value === 'true';
}

function isUnderline(value: string | undefined): boolean {
  switch (value) {
    case 'dash':
    case 'dashHeavy':
    case 'dashLong':
    case 'dashLongHeavy':
    case 'dbl':
    case 'dotDash':
    case 'dotDashHeavy':
    case 'dotDotDash':
    case 'dotDotDashHeavy':
    case 'dotted':
    case 'dottedHeavy':
    case 'heavy':
    case 'sng':
    case 'wavy':
    case 'wavyDbl':
    case 'wavyHeavy':
    case 'words':
      return true;
    default:
      return false;
  }
}

function resolveThemeTypeface(
  typeface: string,
  fontSchemeNode: XmlLookupValue | undefined,
): string {
  let path: readonly [string, string];
  switch (typeface) {
    case '+mj-cs':
      path = ['a:majorFont', 'a:cs'];
      break;
    case '+mj-ea':
      path = ['a:majorFont', 'a:ea'];
      break;
    case '+mj-lt':
      path = ['a:majorFont', 'a:latin'];
      break;
    case '+mn-cs':
      path = ['a:minorFont', 'a:cs'];
      break;
    case '+mn-ea':
      path = ['a:minorFont', 'a:ea'];
      break;
    case '+mn-lt':
      path = ['a:minorFont', 'a:latin'];
      break;
    default:
      return typeface.slice(1);
  }

  const resolved = getTextByPathList(fontSchemeNode, [
    path[0],
    path[1],
    'attrs',
    'typeface',
  ]);
  return resolved || typeface.slice(1);
}

function pushStyleNode(
  styleNodes: XmlLookupValue[],
  styleNode: XmlLookupValue | undefined,
) {
  if (styleNode) styleNodes.push(styleNode);
}

function getLevelPath(lvl: number | string) {
  return `a:lvl${lvl}pPr`;
}

function appendTextBodyStyleNodes(
  styleNodes: XmlLookupValue[],
  textBodyNode: XmlLookupValue | undefined,
  lvl: number | string,
) {
  const lvlPath = getLevelPath(lvl);
  pushStyleNode(
    styleNodes,
    getTextByPathList(textBodyNode, ['a:lstStyle', lvlPath, 'a:defRPr']),
  );
}

function appendShapeStyleNodes(
  styleNodes: XmlLookupValue[],
  shapeNode: XmlLookupValue | undefined,
  lvl: number | string,
) {
  const lvlPath = getLevelPath(lvl);
  pushStyleNode(
    styleNodes,
    getTextByPathList(shapeNode, [
      'p:txBody',
      'a:lstStyle',
      lvlPath,
      'a:defRPr',
    ]),
  );
  pushStyleNode(
    styleNodes,
    getTextByPathList(shapeNode, ['p:txBody', 'a:p', 'a:pPr', 'a:defRPr']),
  );
}

function appendMasterTextStyleNodes(
  styleNodes: XmlLookupValue[],
  type: string,
  lvl: number | string,
  slideMasterTextStyles: XmlLookupValue | undefined,
) {
  const lvlPath = getLevelPath(lvl);

  if (type === 'title' || type === 'ctrTitle' || type === 'subTitle') {
    pushStyleNode(
      styleNodes,
      getTextByPathList(slideMasterTextStyles, [
        'p:titleStyle',
        lvlPath,
        'a:defRPr',
      ]),
    );
    if (type === 'subTitle') {
      pushStyleNode(
        styleNodes,
        getTextByPathList(slideMasterTextStyles, [
          'p:bodyStyle',
          lvlPath,
          'a:defRPr',
        ]),
      );
    }
  } else if (type === 'body') {
    pushStyleNode(
      styleNodes,
      getTextByPathList(slideMasterTextStyles, [
        'p:bodyStyle',
        lvlPath,
        'a:defRPr',
      ]),
    );
  } else {
    pushStyleNode(
      styleNodes,
      getTextByPathList(slideMasterTextStyles, [
        'p:otherStyle',
        lvlPath,
        'a:defRPr',
      ]),
    );
  }
}

function appendDefaultTextStyleNodes(
  styleNodes: XmlLookupValue[],
  lvl: number | string,
  defaultTextStyle: XmlLookupValue | undefined,
) {
  const lvlPath = getLevelPath(lvl);
  pushStyleNode(
    styleNodes,
    getTextByPathList(defaultTextStyle, [lvlPath, 'a:defRPr']),
  );
  pushStyleNode(
    styleNodes,
    getTextByPathList(defaultTextStyle, ['a:defPPr', 'a:defRPr']),
  );
}

function getBaseFontStyleNodes(
  node: XmlLookupValue,
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  lvl: number | string,
) {
  const styleNodes: XmlLookupValue[] = [];
  const runStyleNode = getTextByPathList(node, ['a:rPr']);

  pushStyleNode(styleNodes, runStyleNode);
  if (!runStyleNode) {
    pushStyleNode(styleNodes, getTextByPathList(pNode, ['a:endParaRPr']));
  }
  pushStyleNode(styleNodes, getTextByPathList(pNode, ['a:pPr', 'a:defRPr']));

  appendTextBodyStyleNodes(styleNodes, textBodyNode, lvl);
  appendShapeStyleNodes(styleNodes, slideLayoutSpNode, lvl);
  appendShapeStyleNodes(styleNodes, slideMasterSpNode, lvl);

  return styleNodes;
}

function getFontStyleNodes(
  node: XmlLookupValue,
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  lvl: number | string,
) {
  const styleNodes = getBaseFontStyleNodes(
    node,
    pNode,
    textBodyNode,
    slideLayoutSpNode,
    slideMasterSpNode,
    lvl,
  );
  appendMasterTextStyleNodes(styleNodes, type, lvl, slideMasterTextStyles);

  return styleNodes;
}

function getFontAttr(
  styleNodes: XmlLookupValue[],
  attrName: string,
): string | undefined {
  for (const styleNode of styleNodes) {
    const attrValue = getTextByPathList(styleNode, ['attrs', attrName]);
    if (attrValue) return attrValue;
  }

  return undefined;
}

function getFontTypeface(styleNodes: XmlLookupValue[]): string {
  for (const styleNode of styleNodes) {
    const typeface =
      getTextByPathList(styleNode, ['a:latin', 'attrs', 'typeface']) ||
      getTextByPathList(styleNode, ['a:ea', 'attrs', 'typeface']) ||
      getTextByPathList(styleNode, ['a:cs', 'attrs', 'typeface']) ||
      getTextByPathList(styleNode, ['a:sym', 'attrs', 'typeface']);
    if (typeface) return typeface;
  }

  return '';
}

function getColorFromNode(node: XmlLookupValue, warpObj: PptxParserContext) {
  const solid = getTextByPathList<XmlLookupValue>(node, ['a:solidFill']);
  if (solid) return getSolidFill(solid, undefined, undefined, warpObj);

  const gradient = getTextByPathList<XmlLookupValue>(node, ['a:gradFill']);
  if (gradient) return getGradientFill(gradient, warpObj);

  return '';
}

function getFontColorFromStyleNodes(
  styleNodes: XmlLookupValue[],
  warpObj: PptxParserContext,
) {
  for (const styleNode of styleNodes) {
    const color = getColorFromNode(styleNode, warpObj);
    if (color) return color;
  }

  return '';
}

function getTextShadowFromStyleNodes(
  styleNodes: XmlLookupValue[],
  warpObj: PptxParserContext,
) {
  for (const styleNode of styleNodes) {
    const txtShadow = getTextByPathList(styleNode, [
      'a:effectLst',
      'a:outerShdw',
    ]);
    if (!txtShadow) continue;

    return getShadow(txtShadow, warpObj);
  }

  return null;
}

export function getFontType(
  node: XmlLookupValue,
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  lvl: number | string,
  warpObj: PptxParserContext,
) {
  const styleNodes = getFontStyleNodes(
    node,
    pNode,
    textBodyNode,
    slideLayoutSpNode,
    slideMasterSpNode,
    type,
    slideMasterTextStyles,
    lvl,
  );
  let typeface = getFontTypeface(styleNodes);

  if (!typeface || typeface.startsWith('+')) {
    const fontSchemeNode = getTextByPathList(warpObj['themeContent'], [
      'a:theme',
      'a:themeElements',
      'a:fontScheme',
    ]);

    if (typeface.startsWith('+')) {
      return resolveThemeTypeface(typeface, fontSchemeNode);
    }

    if (type === 'title' || type === 'subTitle' || type === 'ctrTitle') {
      typeface =
        getTextByPathList(fontSchemeNode, [
          'a:majorFont',
          'a:latin',
          'attrs',
          'typeface',
        ]) ||
        getTextByPathList(fontSchemeNode, [
          'a:majorFont',
          'a:ea',
          'attrs',
          'typeface',
        ]) ||
        getTextByPathList(fontSchemeNode, [
          'a:majorFont',
          'a:cs',
          'attrs',
          'typeface',
        ]);
    } else {
      typeface =
        getTextByPathList(fontSchemeNode, [
          'a:minorFont',
          'a:latin',
          'attrs',
          'typeface',
        ]) ||
        getTextByPathList(fontSchemeNode, [
          'a:minorFont',
          'a:ea',
          'attrs',
          'typeface',
        ]) ||
        getTextByPathList(fontSchemeNode, [
          'a:minorFont',
          'a:cs',
          'attrs',
          'typeface',
        ]);
    }
  }

  return typeface ?? '';
}

export function getFontColor(
  node: XmlLookupValue,
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  lvl: number | string,
  pFontStyle: XmlLookupValue | undefined,
  warpObj: PptxParserContext,
) {
  const styleNodes = getBaseFontStyleNodes(
    node,
    pNode,
    textBodyNode,
    slideLayoutSpNode,
    slideMasterSpNode,
    lvl,
  );
  let color = getFontColorFromStyleNodes(styleNodes, warpObj);
  if (color) return color;

  const fallbackReferences = [
    pFontStyle,
    getTextByPathList(slideLayoutSpNode, ['p:style', 'a:fontRef']),
    getTextByPathList(slideMasterSpNode, ['p:style', 'a:fontRef']),
  ];
  for (const reference of fallbackReferences) {
    color = getSolidFill(reference, undefined, undefined, warpObj);
    if (color) return color;
  }

  appendMasterTextStyleNodes(styleNodes, type, lvl, slideMasterTextStyles);
  return getFontColorFromStyleNodes(styleNodes, warpObj);
}

export function getFontSize(
  node: XmlLookupValue,
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  lvl: number | string,
  defaultTextStyle: XmlLookupValue | undefined,
) {
  const styleNodes = getFontStyleNodes(
    node,
    pNode,
    textBodyNode,
    slideLayoutSpNode,
    slideMasterSpNode,
    type,
    slideMasterTextStyles,
    lvl,
  );
  appendDefaultTextStyleNodes(styleNodes, lvl, defaultTextStyle);
  const sz = getFontAttr(styleNodes, 'sz');
  const sizeInHundredths = parseInteger(sz);
  const fontSize =
    sizeInHundredths !== undefined && sizeInHundredths > 0
      ? sizeInHundredths / 100
      : type === 'dt' || type === 'sldNum'
        ? 12
        : 18;

  return fontSize + 'pt';
}

export function getFontBold(
  node: XmlLookupValue,
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  lvl: number | string,
) {
  const styleNodes = getFontStyleNodes(
    node,
    pNode,
    textBodyNode,
    slideLayoutSpNode,
    slideMasterSpNode,
    type,
    slideMasterTextStyles,
    lvl,
  );
  return isDrawingMlTrue(getFontAttr(styleNodes, 'b')) ? 'bold' : '';
}

export function getFontItalic(
  node: XmlLookupValue,
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  lvl: number | string,
) {
  const styleNodes = getFontStyleNodes(
    node,
    pNode,
    textBodyNode,
    slideLayoutSpNode,
    slideMasterSpNode,
    type,
    slideMasterTextStyles,
    lvl,
  );
  return isDrawingMlTrue(getFontAttr(styleNodes, 'i')) ? 'italic' : '';
}

export function getFontDecoration(
  node: XmlLookupValue,
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  lvl: number | string,
) {
  const styleNodes = getFontStyleNodes(
    node,
    pNode,
    textBodyNode,
    slideLayoutSpNode,
    slideMasterSpNode,
    type,
    slideMasterTextStyles,
    lvl,
  );
  return isUnderline(getFontAttr(styleNodes, 'u')) ? 'underline' : '';
}

export function getFontDecorationLine(
  node: XmlLookupValue,
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  lvl: number | string,
) {
  const styleNodes = getFontStyleNodes(
    node,
    pNode,
    textBodyNode,
    slideLayoutSpNode,
    slideMasterSpNode,
    type,
    slideMasterTextStyles,
    lvl,
  );
  const strike = getFontAttr(styleNodes, 'strike');
  return strike === 'sngStrike' || strike === 'dblStrike' ? 'line-through' : '';
}

export function getFontSpace(
  node: XmlLookupValue,
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  lvl: number | string,
) {
  const styleNodes = getFontStyleNodes(
    node,
    pNode,
    textBodyNode,
    slideLayoutSpNode,
    slideMasterSpNode,
    type,
    slideMasterTextStyles,
    lvl,
  );
  const spc = getFontAttr(styleNodes, 'spc');
  const spacing = parseInteger(spc);
  return spacing !== undefined && spacing !== 0 ? spacing / 100 + 'pt' : '';
}

export function getFontSubscript(
  node: XmlLookupValue,
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  lvl: number | string,
) {
  const styleNodes = getFontStyleNodes(
    node,
    pNode,
    textBodyNode,
    slideLayoutSpNode,
    slideMasterSpNode,
    type,
    slideMasterTextStyles,
    lvl,
  );
  const baseline = getFontAttr(styleNodes, 'baseline');
  const offset = parseInteger(baseline);
  if (offset === undefined) return '';
  if (offset > 0) return 'super';
  if (offset < 0) return 'sub';
  return '';
}

export function getFontShadow(
  node: XmlLookupValue,
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  lvl: number | string,
  warpObj: PptxParserContext,
) {
  const styleNodes = getFontStyleNodes(
    node,
    pNode,
    textBodyNode,
    slideLayoutSpNode,
    slideMasterSpNode,
    type,
    slideMasterTextStyles,
    lvl,
  );
  const shadow = getTextShadowFromStyleNodes(styleNodes, warpObj);
  if (!shadow) return '';

  const { h, v, blur, color } = shadow;
  const components = [`${h}pt`, `${v}pt`];
  if (blur) components.push(`${blur}pt`);
  if (color) components.push(color);
  return components.join(' ');
}

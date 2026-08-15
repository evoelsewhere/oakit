import type { XmlLookupValue } from '../../../common';
import type { PptxParserContext } from './context';

import { getTextByPathList } from '../../../common';
import { getShadow } from './shadow';
import { getFillType, getGradientFill, getSolidFill } from './fill';

const UNDERLINE_VALUES = new Set([
  'dash',
  'dashHeavy',
  'dashLong',
  'dashLongHeavy',
  'dbl',
  'dotDash',
  'dotDashHeavy',
  'dotDotDash',
  'dotDotDashHeavy',
  'dotted',
  'dottedHeavy',
  'heavy',
  'sng',
  'wavy',
  'wavyDbl',
  'wavyHeavy',
  'words',
]);

const THEME_FONT_PATHS: Readonly<Record<string, readonly [string, string]>> = {
  '+mj-cs': ['a:majorFont', 'a:cs'],
  '+mj-ea': ['a:majorFont', 'a:ea'],
  '+mj-lt': ['a:majorFont', 'a:latin'],
  '+mn-cs': ['a:minorFont', 'a:cs'],
  '+mn-ea': ['a:minorFont', 'a:ea'],
  '+mn-lt': ['a:minorFont', 'a:latin'],
};

function parseInteger(value: string): number | undefined {
  if (!/^[+-]?\d+$/.test(value)) return undefined;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : undefined;
}

function isDrawingMlTrue(value: string): boolean {
  return value === '1' || value === 'true';
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
  if (!textBodyNode) return;

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
  if (!shapeNode) return;

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
  if (!slideMasterTextStyles) return;

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
  if (!defaultTextStyle) return;

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

function getFontAttr(styleNodes: XmlLookupValue[], attrName: string): string {
  for (const styleNode of styleNodes) {
    const attrValue = getTextByPathList(styleNode, ['attrs', attrName]);
    if (attrValue !== undefined && attrValue !== '') return attrValue;
  }

  return '';
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
  if (!node) return '';

  const fillType = getFillType(node);
  if (fillType === 'SOLID_FILL') {
    return getSolidFill(node['a:solidFill'], undefined, undefined, warpObj);
  }
  if (fillType === 'GRADIENT_FILL') {
    const gradient = getTextByPathList<XmlLookupValue>(node, ['a:gradFill']);
    return gradient ? getGradientFill(gradient, warpObj) : '';
  }

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

    const shadow = getShadow(txtShadow, warpObj);
    if (shadow) return shadow;
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

    if (fontSchemeNode && typeface?.startsWith('+')) {
      const path = THEME_FONT_PATHS[typeface];
      if (!path) return typeface.slice(1);
      return getTextByPathList(fontSchemeNode, [
        path[0],
        path[1],
        'attrs',
        'typeface',
      ]);
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

  return typeface || '';
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

  if (!color) {
    if (pFontStyle)
      color = getSolidFill(pFontStyle, undefined, undefined, warpObj);
    if (!color) {
      const layoutFontStyle = getTextByPathList(slideLayoutSpNode, [
        'p:style',
        'a:fontRef',
      ]);
      if (layoutFontStyle)
        color = getSolidFill(layoutFontStyle, undefined, undefined, warpObj);
    }
    if (!color) {
      const masterFontStyle = getTextByPathList(slideMasterSpNode, [
        'p:style',
        'a:fontRef',
      ]);
      if (masterFontStyle)
        color = getSolidFill(masterFontStyle, undefined, undefined, warpObj);
    }
  }

  if (!color) {
    appendMasterTextStyleNodes(styleNodes, type, lvl, slideMasterTextStyles);
    color = getFontColorFromStyleNodes(styleNodes, warpObj);
  }

  return color || '';
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
  let fontSize =
    sizeInHundredths !== undefined && sizeInHundredths > 0
      ? sizeInHundredths / 100
      : undefined;

  if (
    (!Number.isFinite(fontSize) || !fontSize) &&
    (type === 'dt' || type === 'sldNum')
  )
    fontSize = 12;

  fontSize = !Number.isFinite(fontSize) || !fontSize ? 18 : fontSize;

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
  return UNDERLINE_VALUES.has(getFontAttr(styleNodes, 'u')) ? 'underline' : '';
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
  if (offset === undefined || offset === 0) return '';
  return offset > 0 ? 'super' : 'sub';
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
  if (shadow) {
    const { h, v, blur, color } = shadow;
    if (Number.isFinite(v) && Number.isFinite(h)) {
      const components = [`${h}pt`, `${v}pt`];
      if (blur) components.push(`${blur}pt`);
      if (color) components.push(color);
      return components.join(' ');
    }
  }
  return '';
}

import type { XmlLookupValue } from '../../../common';
import type { PptxParserContext } from './context';
import type { AutoFit } from '../types';

import { RATIO_EMUs_Points } from '../../../common/ooxml/units';
import { getTextByPathList, numberToFixed } from '../../../common';

function getParagraphLevel(node: XmlLookupValue) {
  let lvlIdx = 1;
  const lvlNode = getTextByPathList(node, ['a:pPr', 'attrs', 'lvl']);
  if (lvlNode !== undefined) lvlIdx = parseInt(lvlNode) + 1;
  return lvlIdx;
}

interface ParagraphSpacing {
  lineSpacing?: number | string;
  spaceAfter?: string;
  spaceBefore?: string;
}

interface ParagraphIndent {
  marginLeft?: string;
  textIndent?: string;
}

function getAlignFromTextNode(
  node: XmlLookupValue | undefined,
  levelPath: string,
): string {
  if (!node) return '';

  let algn = getTextByPathList<string>(node, [
    'p:txBody',
    'a:lstStyle',
    levelPath,
    'attrs',
    'algn',
  ]);
  if (!algn)
    algn = getTextByPathList<string>(node, [
      'p:txBody',
      'a:p',
      'a:pPr',
      'attrs',
      'algn',
    ]);

  return algn || '';
}

export function getHorizontalAlign(
  node: XmlLookupValue,
  pNode: XmlLookupValue,
  type: string,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  warpObj: PptxParserContext,
) {
  let algn = getTextByPathList<string>(node, ['a:pPr', 'attrs', 'algn']);

  if (!algn)
    algn = getTextByPathList<string>(pNode, [
      'p:txBody',
      'a:p',
      'a:pPr',
      'attrs',
      'algn',
    ]);

  if (!algn) {
    const lvlIdx = getParagraphLevel(node);
    const lvlStr = 'a:lvl' + lvlIdx + 'pPr';

    algn = getAlignFromTextNode(slideLayoutSpNode, lvlStr);
    if (!algn) algn = getAlignFromTextNode(slideMasterSpNode, lvlStr);

    if (
      !algn &&
      (type === 'title' || type === 'ctrTitle' || type === 'subTitle')
    ) {
      algn = getTextByPathList(warpObj, [
        'slideMasterTextStyles',
        'p:titleStyle',
        lvlStr,
        'attrs',
        'algn',
      ]);
      if (!algn && type === 'subTitle') {
        algn = getTextByPathList(warpObj, [
          'slideMasterTextStyles',
          'p:bodyStyle',
          lvlStr,
          'attrs',
          'algn',
        ]);
      }
    } else if (!algn && type === 'body') {
      algn = getTextByPathList(warpObj, [
        'slideMasterTextStyles',
        'p:bodyStyle',
        lvlStr,
        'attrs',
        'algn',
      ]);
    } else if (!algn) {
      algn = getTextByPathList(warpObj, [
        'slideMasterTextStyles',
        'p:otherStyle',
        lvlStr,
        'attrs',
        'algn',
      ]);
    }
  }

  let align = 'left';
  if (algn) {
    switch (algn) {
      case 'l':
        align = 'left';
        break;
      case 'r':
        align = 'right';
        break;
      case 'ctr':
        align = 'center';
        break;
      case 'just':
        align = 'justify';
        break;
      case 'dist':
        align = 'justify';
        break;
      default:
        align = 'inherit';
    }
  }
  return align;
}

export function getVerticalAlign(
  node: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
) {
  let anchor = getTextByPathList<string>(node, [
    'p:txBody',
    'a:bodyPr',
    'attrs',
    'anchor',
  ]);
  if (!anchor) {
    anchor = getTextByPathList<string>(slideLayoutSpNode, [
      'p:txBody',
      'a:bodyPr',
      'attrs',
      'anchor',
    ]);
    if (!anchor) {
      anchor = getTextByPathList<string>(slideMasterSpNode, [
        'p:txBody',
        'a:bodyPr',
        'attrs',
        'anchor',
      ]);
      if (!anchor) anchor = 't';
    }
  }
  return anchor === 'ctr' ? 'mid' : anchor === 'b' ? 'down' : 'up';
}

export function getTextAutoFit(
  node: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
): AutoFit | null {
  function checkBodyPr(
    bodyPr: XmlLookupValue | undefined,
  ): { result: AutoFit | null } | null {
    if (!bodyPr) return null;

    if (bodyPr['a:noAutofit']) return { result: null };
    else if (bodyPr['a:spAutoFit']) return { result: { type: 'shape' } };
    else if (bodyPr['a:normAutofit']) {
      const fontScale = getTextByPathList(bodyPr['a:normAutofit'], [
        'attrs',
        'fontScale',
      ]);
      if (fontScale) {
        const scalePercent = parseInt(fontScale) / 1000;
        return {
          result: {
            type: 'text',
            fontScale: scalePercent,
          },
        };
      }
      return { result: { type: 'text' } };
    }
    return null;
  }

  const nodeCheck = checkBodyPr(
    getTextByPathList(node, ['p:txBody', 'a:bodyPr']),
  );
  if (nodeCheck) return nodeCheck.result;

  const layoutCheck = checkBodyPr(
    getTextByPathList(slideLayoutSpNode, ['p:txBody', 'a:bodyPr']),
  );
  if (layoutCheck) return layoutCheck.result;

  const masterCheck = checkBodyPr(
    getTextByPathList(slideMasterSpNode, ['p:txBody', 'a:bodyPr']),
  );
  if (masterCheck) return masterCheck.result;

  return null;
}

function pushParagraphStyleNode(
  styleNodes: XmlLookupValue[],
  styleNode: XmlLookupValue | undefined,
) {
  if (styleNode) styleNodes.push(styleNode);
}

function appendTextBodyParagraphStyleNodes(
  styleNodes: XmlLookupValue[],
  textBodyNode: XmlLookupValue,
  lvl: number,
) {
  if (!textBodyNode) return;

  const lvlPath = `a:lvl${lvl}pPr`;
  pushParagraphStyleNode(
    styleNodes,
    getTextByPathList(textBodyNode, ['a:lstStyle', lvlPath]),
  );
}

function appendShapeParagraphStyleNodes(
  styleNodes: XmlLookupValue[],
  shapeNode: XmlLookupValue | undefined,
  lvl: number,
) {
  if (!shapeNode) return;

  const lvlPath = `a:lvl${lvl}pPr`;
  pushParagraphStyleNode(
    styleNodes,
    getTextByPathList(shapeNode, ['p:txBody', 'a:lstStyle', lvlPath]),
  );
  pushParagraphStyleNode(
    styleNodes,
    getTextByPathList(shapeNode, ['p:txBody', 'a:p', 'a:pPr']),
  );
}

function appendMasterTextParagraphStyleNodes(
  styleNodes: XmlLookupValue[],
  type: string,
  lvl: number,
  slideMasterTextStyles: XmlLookupValue | undefined,
) {
  if (!slideMasterTextStyles) return;

  const lvlPath = `a:lvl${lvl}pPr`;

  if (type === 'title' || type === 'ctrTitle' || type === 'subTitle') {
    pushParagraphStyleNode(
      styleNodes,
      getTextByPathList(slideMasterTextStyles, ['p:titleStyle', lvlPath]),
    );
    if (type === 'subTitle') {
      pushParagraphStyleNode(
        styleNodes,
        getTextByPathList(slideMasterTextStyles, ['p:bodyStyle', lvlPath]),
      );
    }
  } else if (type === 'body') {
    pushParagraphStyleNode(
      styleNodes,
      getTextByPathList(slideMasterTextStyles, ['p:bodyStyle', lvlPath]),
    );
  } else {
    pushParagraphStyleNode(
      styleNodes,
      getTextByPathList(slideMasterTextStyles, ['p:otherStyle', lvlPath]),
    );
  }
}

function appendDefaultTextParagraphStyleNodes(
  styleNodes: XmlLookupValue[],
  defaultTextStyle: XmlLookupValue | undefined,
  lvl: number,
) {
  if (!defaultTextStyle) return;

  const lvlPath = `a:lvl${lvl}pPr`;
  pushParagraphStyleNode(
    styleNodes,
    getTextByPathList(defaultTextStyle, [lvlPath]),
  );
  pushParagraphStyleNode(
    styleNodes,
    getTextByPathList(defaultTextStyle, ['a:defPPr']),
  );
}

function getParagraphStyleNodes(
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  warpObj: PptxParserContext,
) {
  if (!pNode) return null;

  const pPrNode = pNode['a:pPr'];
  const lvl = getParagraphLevel(pNode);
  const styleNodes: XmlLookupValue[] = [];

  pushParagraphStyleNode(styleNodes, pPrNode);
  appendTextBodyParagraphStyleNodes(styleNodes, textBodyNode, lvl);
  appendShapeParagraphStyleNodes(styleNodes, slideLayoutSpNode, lvl);
  appendShapeParagraphStyleNodes(styleNodes, slideMasterSpNode, lvl);
  appendMasterTextParagraphStyleNodes(
    styleNodes,
    type,
    lvl,
    slideMasterTextStyles,
  );
  appendDefaultTextParagraphStyleNodes(
    styleNodes,
    getTextByPathList(warpObj, ['defaultTextStyle']),
    lvl,
  );

  return styleNodes;
}

function getLineSpacingValue(spacingNode: XmlLookupValue | undefined) {
  const spcPct = getTextByPathList(spacingNode, ['a:spcPct', 'attrs', 'val']);
  const spcPts = getTextByPathList(spacingNode, ['a:spcPts', 'attrs', 'val']);

  if (spcPct) return parseInt(spcPct) / 1000 / 100;
  if (spcPts) return parseInt(spcPts) / 100 + 'pt';

  return undefined;
}

function getParagraphSpacingValue(spacingNode: XmlLookupValue | undefined) {
  const spcPct = getTextByPathList(spacingNode, ['a:spcPct', 'attrs', 'val']);
  const spcPts = getTextByPathList(spacingNode, ['a:spcPts', 'attrs', 'val']);

  if (spcPct) return parseInt(spcPct) / 1000 + 'em';
  if (spcPts) return parseInt(spcPts) / 100 + 'pt';

  return undefined;
}

function getParagraphIndentValue(styleNode: XmlLookupValue, attrName: string) {
  const val = getTextByPathList(styleNode, ['attrs', attrName]);

  if (val !== undefined && val !== '')
    return numberToFixed(parseInt(val) * RATIO_EMUs_Points) + 'pt';

  return undefined;
}

export function getParagraphSpacing(
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  warpObj: PptxParserContext,
) {
  const styleNodes = getParagraphStyleNodes(
    pNode,
    textBodyNode,
    slideLayoutSpNode,
    slideMasterSpNode,
    type,
    slideMasterTextStyles,
    warpObj,
  );
  if (!styleNodes) return null;

  const spacing: ParagraphSpacing = {};

  for (const styleNode of styleNodes) {
    if (spacing.lineSpacing === undefined) {
      const lineSpacing = getLineSpacingValue(styleNode['a:lnSpc']);
      if (lineSpacing !== undefined) spacing.lineSpacing = lineSpacing;
    }

    if (spacing.spaceBefore === undefined) {
      const spaceBefore = getParagraphSpacingValue(styleNode['a:spcBef']);
      if (spaceBefore !== undefined) spacing.spaceBefore = spaceBefore;
    }

    if (spacing.spaceAfter === undefined) {
      const spaceAfter = getParagraphSpacingValue(styleNode['a:spcAft']);
      if (spaceAfter !== undefined) spacing.spaceAfter = spaceAfter;
    }
  }

  return Object.keys(spacing).length > 0 ? spacing : null;
}

export function getParagraphIndent(
  pNode: XmlLookupValue,
  textBodyNode: XmlLookupValue,
  slideLayoutSpNode: XmlLookupValue | undefined,
  slideMasterSpNode: XmlLookupValue | undefined,
  type: string,
  slideMasterTextStyles: XmlLookupValue | undefined,
  warpObj: PptxParserContext,
) {
  const styleNodes = getParagraphStyleNodes(
    pNode,
    textBodyNode,
    slideLayoutSpNode,
    slideMasterSpNode,
    type,
    slideMasterTextStyles,
    warpObj,
  );
  if (!styleNodes) return null;

  const indent: ParagraphIndent = {};

  for (const styleNode of styleNodes) {
    if (indent.marginLeft === undefined) {
      const marginLeft = getParagraphIndentValue(styleNode, 'marL');
      if (marginLeft !== undefined) indent.marginLeft = marginLeft;
    }

    if (indent.textIndent === undefined) {
      const textIndent = getParagraphIndentValue(styleNode, 'indent');
      if (textIndent !== undefined) indent.textIndent = textIndent;
    }
  }

  return Object.keys(indent).length > 0 ? indent : null;
}
